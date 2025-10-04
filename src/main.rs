mod display;
mod image_preprocessor;
mod pdf_document;
mod pdf_stream_writer;

use anyhow::Result;
use display::DisplayMap;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufRead, BufReader, Read, Seek, Write};
use std::time::Instant;
use zip::ZipArchive;

use image_preprocessor::ImagePreprocessor;
use pdf_document::{ImageFormat, PdfDocument};

fn parse_document_xml<R, W>(
    mut archive: ZipArchive<R>,
    pdf_document: &mut PdfDocument<W>,
    media_lookup: Fn(&str) -> Option<String>,
) -> Result<()>
where
    R: BufRead + Seek,
    W: Write + Seek,
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

    let mut current_text = String::new();
    let mut buf = Vec::new();

    let mut in_drawing = false;
    let mut seen_rid = HashSet::new();
    let mut current_page_objs = Vec::new();
    let mut current_page_img_objs = Vec::new();

    let create_image_obj = |rid: &str| {
        if seen_rid.contains(rid) {
            return;
        }

        seen_rid.insert(rid);
        println!("Found new image ref: {}", rid);
        if let Some(path) = media_lookup(&rid) {
            println!("image file path: {}", path);
            let f = archive.by_name(&path)?;
            let image_id = pdf_document.new_img_obj(f, image_format)?;
            current_page_img_objs.push(image_id);
        } else {
            println!("No data found");
        }
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,

            // --- text nodes ---
            Ok(Event::Text(e)) if !in_drawing => {
                let t = e.unescape().unwrap_or_default().to_string();
                current_text.push_str(&t);
                current_text.push(' ');
            }

            // --- paragraph end ---
            Ok(Event::End(e)) if e.name().as_ref() == b"w:p" => {
                if !current_text.trim().is_empty() {
                    println!("Paragraph: {}", current_text.trim());
                    current_page_objs.push(pdf_document.new_text_obj(current_text.trim())?);
                }
                current_text.clear();
            }

            // --- start tags (detect drawings or breaks) ---
            Ok(Event::Start(e)) => {
                let name = e.name().into_inner();

                // Detect page breaks: <w:br w:type="page"/>
                if name == b"w:br" {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"w:type" && a.value.as_ref() == b"page" {
                            println!("📄 Detected manual page break (<w:br w:type=\"page\"/>)");
                            pdf_document
                                .new_page_obj(&current_page_objs, &current_page_img_objs)?;
                            current_page_objs.clear();
                            current_page_img_objs.clear();
                        }
                    }
                }

                // Detect section breaks: <w:sectPr>
                if name == b"w:sectPr" {
                    println!("📄 Detected section break (new page)");
                    pdf_document.new_page_obj(&current_page_objs, &current_page_img_objs)?;
                    current_page_objs.clear();
                    current_page_img_objs.clear();
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
                        if !seen_rid.contains(&rid) {
                            println!("Found new image ref: {}", rid);
                            if let Some(path) = media_lookup(&rid, rel_map) {
                                println!("path: {}", path);
                                let f = archive.by_name(&path)?;
                                let image_format = image_format_from_path(&path);
                                let image_id = pdf_document.new_img_obj(f, image_format)?;
                                current_page_img_objs.push(image_id);
                            } else {
                                println!("No data found");
                            }
                            seen_rid.insert(rid);
                        }
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
                            println!("📄 Detected manual page break (<w:br w:type=\"page\"/>)");
                            pdf_document
                                .new_page_obj(&current_page_objs, &current_page_img_objs)?;
                            current_page_objs.clear();
                            current_page_img_objs.clear();
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
                        if !seen_rid.contains(&rid) {
                            println!("Found new image ref: {}", rid);
                            if let Some(path) = media_lookup(&rid, rel_map) {
                                println!("path: {}", path);
                                let f = archive.by_name(&path)?;
                                let image_format = image_format_from_path(&path);
                                let image_tuple = pdf_document.new_img_obj(f, image_format)?;
                                current_page_img_objs.push(image_tuple);
                            } else {
                                println!("No data found");
                            }
                            seen_rid.insert(rid);
                        }
                    }
                }
            }

            _ => {}
        }
        buf.clear();
    }

    // writes any remaining objects
    pdf_document.new_page_obj(&current_page_objs, &current_page_img_objs)?;
    Ok(())
}

fn image_format_from_path(path: &str) -> ImageFormat {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        ImageFormat::Jpeg
    } else if lower.ends_with(".png") {
        ImageFormat::Png
    } else if lower.ends_with(".gif") {
        ImageFormat::Gif
    } else if lower.ends_with(".bmp") {
        ImageFormat::Bmp
    } else {
        ImageFormat::Unknown
    }
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
    let path = "test_files/large_business_plan.docx";
    let out_path = "test_files/out.pdf";
    let file = File::open(path)?;

    //let mut archive = ZipArchive::new(BufReader::new(file))?;

    let start = Instant::now();
    let rel_map = build_rel_map(&mut archive)?;
    //println!("{}", DisplayMap(&rel_map));

    let image_preprocessor = ImagePreprocessor::preprocess_images(path)?;
    let image_map = &image_preprocessor.image_map;

    //let output_file = File::create(out_path)?;
    //let mut pdf_document = PdfDocument::new(output_file)?;
    //parse_document_xml(archive, &rel_map, &mut pdf_document)?;
    //pdf_document.finish_document()?;

    let media_lookup = |rid: &str| {
        if let Some(target) = rel_map.get(rid) {
            // Normalise to full ZIP path
            let zip_path = if let Some(stripped) = target.strip_prefix("../") {
                format!("word/{}", stripped)
            } else {
                format!("word/{}", target)
            };

            if let Some(tmp_img_path) = image_map.get(&zip_path) {
                let full_path = image_preprocessor.temp_dir.path().join(tmp_img_path);
                Some(full_path.to_string_lossy().to_string())
            } else {
                None
            }
        } else {
            None
        }
    };

    let duration = Instant::now() - start;
    println!("Finished parsing DOCX, took {}ms", duration.as_millis());

    println!("tempdir path:\n{:?}", image_preprocessor.temp_dir.path());

    let ten_secs = std::time::Duration::from_secs(120);
    //std::thread::sleep(ten_secs);
    Ok(())
}
