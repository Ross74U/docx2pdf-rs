use anyhow::Result;
use image::GenericImageView;
use rayon::prelude::*;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::path::PathBuf;
use tempfile::TempDir;
use turbojpeg::{Compressor, Image, PixelFormat};
use zip::ZipArchive;

#[derive(Clone)]
pub struct ImageParams {
    pub w: u32,
    pub h: u32,
    pub len: usize,
    pub path: PathBuf,
}

pub struct ImagePreprocessor {
    pub temp_dir: TempDir,
    pub image_map: HashMap<String, ImageParams>,
}

impl ImagePreprocessor {
    pub fn preprocess_images(docx_path: &str) -> Result<Self> {
        let temp_dir = TempDir::new()?;

        // Step 1: Collect image paths from the archive
        let image_entries = {
            let file = File::open(docx_path)?;
            let mut archive = ZipArchive::new(BufReader::new(file))?;

            let mut entries = Vec::new();
            for i in 0..archive.len() {
                let file = archive.by_index(i)?;
                let name = file.name().to_string();

                if name.starts_with("word/media/")
                    && (name.ends_with(".png")
                        || name.ends_with(".bmp")
                        || name.ends_with(".gif")
                        || name.ends_with(".jpeg")
                        || name.ends_with(".jpg"))
                {
                    entries.push((i, name));
                }
            }
            entries
        };

        println!("Found {} images to process", image_entries.len());

        // Step 2: Process images in parallel, each with its own archive handle
        let image_map: HashMap<String, ImageParams> = image_entries
            .into_par_iter()
            .map(|(index, path)| {
                println!("{:?}", path);

                let file = File::open(docx_path)?;
                let mut archive = ZipArchive::new(BufReader::new(file))?;

                // Process this specific image
                let mut zip_file = archive.by_index(index)?;
                let image_params = Self::process_single_image(&mut zip_file, &path, &temp_dir)?;

                Ok((path, image_params))
            })
            .collect::<Result<HashMap<_, _>>>()?;

        Ok(Self {
            temp_dir,
            image_map,
        })
    }

    fn process_single_image<R: Read>(
        zip_file: &mut R,
        original_path: &str,
        temp_dir: &TempDir,
    ) -> Result<ImageParams> {
        // Generate consistent temp filename
        let filename = original_path
            .strip_prefix("word/media/")
            .unwrap_or(original_path);

        let path = std::path::Path::new(filename);
        let filename_with_jpg = path.with_extension("jpg");
        println!("path: {:?}", filename_with_jpg);
        let temp_path = temp_dir.path().join(filename_with_jpg);

        let mut data = Vec::new();
        zip_file.read_to_end(&mut data)?;
        let len = data.len();

        let start = std::time::Instant::now();
        // Decode it using the `image` crate
        let img = image::load_from_memory(&data)?;

        let (width, height) = img.dimensions();

        // Convert dynamic image to raw RGB8 pixel buffer
        let rgb = img.to_rgb8();
        let pixels: &[u8] = rgb.as_raw(); // this is what TurboJPEG needs

        // Initialize TurboJPEG compressor
        let mut compressor = Compressor::new()?;
        compressor.set_quality(50)?;
        compressor.set_subsamp(turbojpeg::Subsamp::Sub2x1)?; // choose 4:2:2 subsampling

        // Prepare image descriptor for TurboJPEG
        let image = Image {
            pixels,
            width: width as usize,
            height: height as usize,
            pitch: (width as usize) * 3, // 3 bytes per pixel for RGB
            format: PixelFormat::RGB,
        };
        let jpeg_data = compressor.compress_to_vec(image)?;
        std::fs::write(&temp_path, jpeg_data)?;

        Ok(ImageParams {
            path: temp_path,
            w: width,
            h: height,
            len,
        })
    }
}
