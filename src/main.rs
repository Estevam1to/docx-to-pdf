use anyhow::Result;
use log::{error, info};

mod docx_reader;
mod pdf_writer;
mod utils;

use crate::docx_reader::read_docx;
use crate::pdf_writer::convert_paragraphs_to_pdf;

pub const PAGE_WIDTH: f32 = 210.0;
pub const PAGE_HEIGHT: f32 = 297.0;
pub const MARGIN: f32 = 10.0;
pub const LINE_HEIGHT: f32 = 6.0;
pub const PARAGRAPH_SPACING: f32 = 8.0;
pub const FONT_SIZE: f32 = 11.0;

fn main() -> Result<()> {
    env_logger::init();

    let args: Vec<String> = std::env::args().collect();
    if args.len() < 3 {
        anyhow::bail!("Usage: {} <input.docx> <output.pdf>", args[0]);
    }
    let docx_path = &args[1];
    let pdf_path = &args[2];

    info!("Starting conversion from {} to {}", docx_path, pdf_path);

    match convert_docx_to_pdf(docx_path, pdf_path) {
        Ok(_) => {
            info!("Conversion completed successfully");
            Ok(())
        }
        Err(e) => {
            error!("Conversion failed: {:?}", e);
            Err(e)
        }
    }
}

fn convert_docx_to_pdf(docx_path: &str, pdf_path: &str) -> Result<()> {
    let content = read_docx(docx_path)?;
    info!("Successfully read DOCX file. Converting to PDF...");
    convert_paragraphs_to_pdf(content, pdf_path)?;
    Ok(())
}