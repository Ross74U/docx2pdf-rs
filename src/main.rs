mod image_preprocessor;
mod pdf_document;
mod pdf_stream_writer;

use anyhow::Result;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufRead, BufReader, Read, Seek, Write};
use std::time::Instant;
use zip::ZipArchive;

use image_preprocessor::{ImageParams, ImagePreprocessor};
use pdf_document::PdfDocument;

struct ParserEnv<'a, W: Write + Seek> {
    pdf_document: &'a mut PdfDocument<W>,
    current_text: String,
    seen_rid: HashSet<String>,
    current_page_objs: Vec<u32>,
    current_page_img_objs: Vec<(u32, u32)>,
}

fn parse_document_xml<R, W, F>(
    mut archive: ZipArchive<R>,
    mut pdf_document: &mut PdfDocument<W>,
    media_lookup: F,
) -> Result<()>
where
    R: BufRead + Seek,
    W: Write + Seek,
    F: Fn(&str) -> Option<(String, ImageParams)>,
{
    // Load entire document.xml into memory
    let doc_xml = {
        let mut buf = Vec::new();
        let mut doc_xml = archive.by_name("word/document.xml")?;
        doc_xml.read_to_end(&mut buf)?;
        buf
    };
    let mut buf_doc_xml = BufReader::new(&doc_xml[..]);
    let mut reader = Reader::from_reader(&mut buf_doc_xml);

    let mut buf = Vec::new();
    let mut in_drawing = false;

    let mut env = ParserEnv {
        pdf_document: &mut pdf_document,
        current_text: String::new(),
        seen_rid: HashSet::new(),
        current_page_objs: Vec::new(),
        current_page_img_objs: Vec::new(),
    };

    let mut _create_image_obj = |env: &mut ParserEnv<W>, rid: String| {
        if env.seen_rid.contains(&rid) {
            return;
        }
        if let Some((path, params)) = media_lookup(&rid) {
            println!("image file path: {}", path);
            if let Ok(f) = File::open(&path) {
                if let Ok(image_id) = env
                    .pdf_document
                    .new_img_obj(f, params.w, params.h, params.len)
                {
                    env.current_page_img_objs.push(image_id);
                } else {
                    eprintln!("Image object could not be created, skipping");
                }
            } else {
                eprintln!("Image file could not be opened, skipping");
            }
        } else {
            eprintln!("No data found, skipping");
        }
        env.seen_rid.insert(rid);
    };

    let mut _create_text_obj = |env: &mut ParserEnv<W>| {
        if !env.current_text.trim().is_empty() {
            println!("Paragraph: {}", env.current_text.trim());
            if let Ok(obj) = env.pdf_document.new_text_obj(env.current_text.trim()) {
                env.current_page_objs.push(obj);
            } else {
                eprintln!("Text object could not be created, skipping");
            }
        }
        env.current_text.clear();
    };

    let mut _create_page_obj = |env: &mut ParserEnv<W>| {
        if env
            .pdf_document
            .new_page_obj(&env.current_page_objs, &env.current_page_img_objs)
            .is_err()
        {
            eprintln!("Page object could not be created, skipping");
        }
        env.current_page_objs.clear();
        env.current_page_img_objs.clear();
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,

            // --- text nodes ---
            Ok(Event::Text(e)) if !in_drawing => {
                let t = e.unescape().unwrap_or_default().to_string();
                env.current_text.push_str(&t);
                env.current_text.push(' ');
            }

            // --- paragraph end ---
            Ok(Event::End(e)) if e.name().as_ref() == b"w:p" => {
                _create_text_obj(&mut env);
            }

            // --- start tags (detect drawings or breaks) ---
            Ok(Event::Start(e)) => {
                let name = e.name().into_inner();

                // Detect page breaks: <w:br w:type="page"/>
                if name == b"w:br" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"w:type" && a.value.as_ref() == b"page" {
                            _create_page_obj(&mut env);
                        }
                    }
                }

                // Detect section breaks: <w:sectPr>
                if name == b"w:sectPr" {
                    _create_page_obj(&mut env);
                }

                // Handle drawings/images
                if name == b"w:drawing" || name == b"wp:inline" || name == b"wp:extent" {
                    in_drawing = true;
                }

                if name.ends_with(b"blip") {
                    if let Some(attr) = e
                        .attributes()
                        .with_checks(false)
                        .filter_map(|a| a.ok())
                        .find(|a| a.key.as_ref().ends_with(b"embed"))
                    {
                        let rid = String::from_utf8_lossy(&attr.value).to_string();
                        _create_image_obj(&mut env, rid);
                    }
                }
            }

            Ok(Event::End(e)) => {
                let name = e.name().into_inner();
                if name == b"w:drawing" || name == b"wp:inline" || name == b"wp:extent" {
                    in_drawing = false;
                }
            }

            Ok(Event::Empty(e)) => {
                let name = e.name().into_inner();

                // manual page break (self-closing)
                if name == b"w:br" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"w:type" && a.value.as_ref() == b"page" {
                            _create_page_obj(&mut env);
                        }
                    }
                }

                // inline image reference
                if name.ends_with(b"blip") {
                    if let Some(attr) = e
                        .attributes()
                        .with_checks(false)
                        .filter_map(|a| a.ok())
                        .find(|a| a.key.as_ref().ends_with(b"embed"))
                    {
                        let rid = String::from_utf8_lossy(&attr.value).to_string();
                        _create_image_obj(&mut env, rid);
                    }
                }
            }

            _ => {}
        }
        buf.clear();
    }

    // writes any remaining objects
    if env.current_page_img_objs.len() > 0 || env.current_page_objs.len() > 0 {
        env.pdf_document
            .new_page_obj(&env.current_page_objs, &env.current_page_img_objs)?;
    }

    Ok(())
}

fn build_rel_map<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<HashMap<String, String>> {
    let mut rels_map = HashMap::new();

    // It's always at this path relative to word/document.xml
    let mut rels_file = archive.by_name("word/_rels/document.xml.rels")?;
    let mut xml = String::new();
    std::io::Read::read_to_string(&mut rels_file, &mut xml)?;

    let mut reader = Reader::from_str(&xml);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if e.name().as_ref().ends_with(b"Relationship") =>
            {
                // extract Id and Target attributes
                let mut id = None;
                let mut target = None;

                for attr in e.attributes().with_checks(false).flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = Some(String::from_utf8_lossy(&attr.value).to_string()),
                        b"Target" => {
                            target = Some(String::from_utf8_lossy(&attr.value).to_string())
                        }
                        _ => {}
                    }
                }

                if let (Some(i), Some(t)) = (id, target) {
                    rels_map.insert(i, t);
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(rels_map)
}

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    let (out_path, path) = if args.len() > 3 && args[1] == "-o" {
        (args[2].clone(), args[3].clone())
    } else {
        panic!("Invalid args, should be formatted in './program -o [output_path] [input_path]'");
    };

    let file = File::open(&path)?;
    let mut archive = ZipArchive::new(BufReader::new(file))?;

    let start = Instant::now();
    let rel_map = build_rel_map(&mut archive)?;

    let image_preprocessor = ImagePreprocessor::preprocess_images(&path)?;
    let image_map = image_preprocessor.image_map;

    let media_lookup = |rid: &str| {
        if let Some(target) = rel_map.get(rid) {
            // Normalise to full ZIP path
            let zip_path = if let Some(stripped) = target.strip_prefix("../") {
                format!("word/{}", stripped)
            } else {
                format!("word/{}", target)
            };

            if let Some(image_params) = image_map.get(&zip_path) {
                let tmp_img_path = image_params.path.clone();
                let full_path = image_preprocessor
                    .temp_dir
                    .path()
                    .join(tmp_img_path)
                    .to_string_lossy()
                    .to_string();

                Some((full_path, image_params.clone()))
            } else {
                None
            }
        } else {
            None
        }
    };

    let output_file = File::create(out_path)?;
    let mut pdf_document = PdfDocument::new(output_file)?;
    parse_document_xml(archive, &mut pdf_document, &media_lookup)?;
    pdf_document.finish_document()?;

    let duration = Instant::now() - start;
    println!("Finished parsing DOCX, took {}ms", duration.as_millis());
    Ok(())
}
