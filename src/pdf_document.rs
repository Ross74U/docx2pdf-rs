use super::pdf_stream_writer::PdfStreamWriter;
use anyhow::Result;
use std::io::{Cursor, Read, Seek, Write};

pub struct PdfDocument<W>
where
    W: Write + Seek,
{
    writer: PdfStreamWriter<W>,
    pages_id: u32,      // reserved object id for /Pages
    page_ids: Vec<u32>, // vector of page object ids used to build /Pages
}

impl<W> PdfDocument<W>
where
    W: Write + Seek,
{
    pub fn new(sink: W) -> Result<PdfDocument<W>> {
        let mut writer = PdfStreamWriter::new(sink)?;
        let pages_id = writer.reserve_object()?;
        Ok(PdfDocument {
            writer,
            pages_id,
            page_ids: vec![],
        })
    }

    /// writes a new unused text stream object, returning the object_id
    pub fn new_text_obj(&mut self, text: &str) -> Result<u32> {
        let stream_content = format!(
            "BT /F1 12 Tf 0 720 Td ({}) Tj ET",
            text.replace('(', "\\(").replace(')', "\\)")
        );
        let content_bytes = stream_content.as_bytes();
        let length = content_bytes.len();

        let mut buf = Vec::new();
        writeln!(buf, "<< /Length {} >>", length)?;
        writeln!(buf, "stream")?;
        buf.extend_from_slice(content_bytes);
        writeln!(buf, "\nendstream")?;

        // Then call your low-level function
        let reader = Cursor::new(buf);
        let obj_id = self.writer.write_object(reader)?;
        Ok(obj_id)
    }

    /// assumes the image data is of JPEG
    pub fn new_img_obj<R: Read>(
        &mut self,
        mut image_stream: R,
        w: u32,
        h: u32,
        len: u64,
    ) -> Result<(u32, u32)> {
        // For JPEG, stream directly without re-encoding
        // Read just enough to get dimensions

        // We need to know the length, so collect the data
        let mut image_data = Vec::new();
        image_stream.read_to_end(&mut image_data)?;

        // Now write with known length
        let img_dict = format!(
            "<< /Type /XObject /Subtype /Image \
               /Width {} /Height {} \
               /ColorSpace /DeviceRGB \
               /BitsPerComponent 8 \
               /Filter /DCTDecode \
               /Length {} >>\nstream\n",
            w,
            h,
            image_data.len()
        );

        let mut composed = Vec::with_capacity(final_dict.len() + image_data.len() + 20);
        composed.extend_from_slice(final_dict.as_bytes());
        composed.extend_from_slice(&image_data);
        composed.extend_from_slice(b"\nendstream");

        let image_obj_id = self.writer.write_object(Cursor::new(composed))?;

        // Create simple content stream
        let content = format!("q\n500 0 0 500 0 0 cm\n/Im{} Do\nQ\n", image_obj_id);
        let content_stream = format!(
            "<< /Length {} >>\nstream\n{}\nendstream",
            content.len(),
            content
        );

        let content_stream_id = self
            .writer
            .write_object(Cursor::new(content_stream.into_bytes()))?;

        return Ok((image_obj_id, content_stream_id));
    }

    pub fn new_page_obj(
        &mut self,
        current_page_objs: &[u32],            // Text content stream IDs
        current_page_img_objs: &[(u32, u32)], // Vec of (image_obj_id, content_stream_id) tuples
    ) -> Result<u32> {
        // Collect all content stream IDs (text + image drawing commands)
        let mut all_content_streams = Vec::new();

        // Add text content streams
        all_content_streams.extend_from_slice(current_page_objs);

        // Add image content streams and build resource dictionary
        let mut xobject_entries = Vec::new();
        for (image_obj_id, content_stream_id) in current_page_img_objs {
            all_content_streams.push(*content_stream_id);

            // Use image_obj_id as the resource name suffix for uniqueness
            let resource_name = format!("Im{}", image_obj_id);
            xobject_entries.push((resource_name, *image_obj_id));
        }

        // Build Contents array (all content streams)
        let contents = all_content_streams
            .iter()
            .map(|id| format!("{} 0 R", id))
            .collect::<Vec<_>>()
            .join(" ");

        // Build Resources dictionary if we have images
        let resources = if !xobject_entries.is_empty() {
            let mut xobject_dict = String::from("<< ");
            for (resource_name, img_obj_id) in xobject_entries {
                xobject_dict.push_str(&format!("/{} {} 0 R ", resource_name, img_obj_id));
            }
            xobject_dict.push_str(">>");

            format!("/Resources << /XObject {} >>", xobject_dict)
        } else {
            String::new()
        };

        // Create the page object
        let page_dict = format!(
            "<< /Type /Page \
           /Parent {} 0 R \
           /MediaBox [0 0 595 842] \
           {} \
           /Contents [{}] >>",
            self.pages_id, resources, contents
        );

        let reader = Cursor::new(page_dict.into_bytes());
        let page_id = self.writer.write_object(reader)?;
        self.page_ids.push(page_id);

        Ok(page_id)
    }

    pub fn finish_document(mut self) -> Result<()> {
        // 1️⃣ Create the /Pages object listing all page IDs
        let kids = self.page_ids.iter().fold(String::new(), |mut acc, &id| {
            use std::fmt::Write as _;
            let _ = write!(acc, "{} 0 R ", id);
            acc
        });

        let pages_dict = format!(
            "<< /Type /Pages /Count {} /Kids [{}] >>",
            self.page_ids.len(),
            kids
        );

        // Ensure the reserved `pages_id` is actually used
        assert_eq!(
            self.pages_id, 1,
            "expected pages to be object 1 (just convention)"
        );
        let reader = Cursor::new(pages_dict.into_bytes());
        let pages_id = self.writer.write_object(reader)?;

        // 2️⃣ Create the /Catalog object pointing to /Pages
        let catalog_dict = format!("<< /Type /Catalog /Pages {} 0 R >>", pages_id);
        let reader = Cursor::new(catalog_dict.into_bytes());
        let catalog_id = self.writer.write_object(reader)?;

        self.writer.finish(catalog_id)
    }
}

fn parse_jpeg_size(buf: &[u8]) -> Option<(u32, u32)> {
    let mut i = 2;
    while i + 9 < buf.len() {
        if buf[i] != 0xFF {
            i += 1;
            continue;
        }
        let marker = buf[i + 1];
        if (0xC0..=0xC3).contains(&marker) || (0xC5..=0xC7).contains(&marker) {
            let h = u16::from_be_bytes([buf[i + 5], buf[i + 6]]) as u32;
            let w = u16::from_be_bytes([buf[i + 7], buf[i + 8]]) as u32;
            return Some((w, h));
        }
        let segment_len = u16::from_be_bytes([buf[i + 2], buf[i + 3]]) as usize;
        i += 2 + segment_len;
    }
    None
}
