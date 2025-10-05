use anyhow::Result;
use std::io::{Read, Seek, Write};

pub struct PdfStreamWriter<W: Write + Seek> {
    sink: W,
    offsets: Vec<u64>,
    pub next_obj_id: u32,
}

impl<W: Write + Seek> PdfStreamWriter<W> {
    pub fn new(mut sink: W) -> Result<Self> {
        sink.write_all(b"%PDF-1.7\n")?;
        Ok(Self {
            sink,
            offsets: vec![],
            next_obj_id: 1,
        })
    }

    pub fn write_object_with<F>(&mut self, writer: &mut F) -> Result<u32>
    where
        F: FnMut(&mut W) -> Result<()>,
    {
        let id = self._new_object()?;
        writer(&mut self.sink)?;
        write!(self.sink, "\n");
        self._finish_object()?;
        Ok(id)
    }

    /// Stream arbitrary bytes from `reader` into the PDF output as an object.
    /// returns the id of the object
    ///
    /// This just copies all bytes until EOF. It's agnostic about the
    /// structure of the PDF object — the caller decides what to write
    /// before and after. Ideal for large, already‑encoded object data
    /// such as content streams, images, or dictionary bodies generated
    /// elsewhere.
    pub fn write_object<R: Read>(&mut self, mut reader: R) -> Result<u32> {
        let id = self._new_object()?;
        let mut buf = [0u8; 8192];
        loop {
            let n = reader.read(&mut buf)?;
            if n == 0 {
                break; // EOF
            }
            self.sink.write_all(&buf[..n])?;
        }
        self._finish_object()?;
        Ok(id)
    }

    // Completes a pdf file
    // writes the xref table and EOF
    pub fn finish(mut self, root_id: u32) -> Result<()> {
        // write xref table at current position
        let xref_start = self.sink.stream_position()?;
        writeln!(self.sink, "xref")?;
        writeln!(self.sink, "0 {}", self.offsets.len() + 1)?;
        writeln!(self.sink, "0000000000 65535 f ")?;
        for &offset in &self.offsets {
            writeln!(self.sink, "{:010} 00000 n ", offset)?;
        }
        writeln!(
            self.sink,
            "trailer << /Size {} /Root {} 0 R >>",
            self.offsets.len() + 1,
            root_id
        )?;
        writeln!(self.sink, "startxref")?;
        writeln!(self.sink, "{}", xref_start)?;
        writeln!(self.sink, "%%EOF")?;
        Ok(())
    }

    pub fn reserve_object(&mut self) -> Result<u32> {
        let object_id = self.next_obj_id;
        self.next_obj_id += 1;
        Ok(object_id)
    }

    pub fn write_object_with_reserved_id<F>(&mut self, id: u32, writer: &mut F) -> Result<()>
    where
        F: FnMut(&mut W) -> Result<()>,
    {
        let pos = self.sink.stream_position()?;
        self.offsets.push(pos);
        writeln!(self.sink, "{} 0 obj", id)?;
        writer(&mut self.sink)?;
        writeln!(self.sink, "\nendobj")?;
        Ok(())
    }

    fn _new_object(&mut self) -> Result<u32> {
        let id = self.next_obj_id;
        self.next_obj_id += 1;
        let pos = self.sink.stream_position()?;
        self.offsets.push(pos);
        writeln!(self.sink, "{} 0 obj", id)?;
        Ok(id)
    }

    fn _finish_object(&mut self) -> Result<()> {
        writeln!(self.sink, "\nendobj")?;
        Ok(())
    }
}
