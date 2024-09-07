use printpdf::image_crate::codecs::jpeg::JpegDecoder as PrintPdfJpegDecoder;
use printpdf::image_crate::codecs::png::PngDecoder as PrintPdfPngDecoder;
use printpdf::image_crate::{guess_format, ImageFormat};
use printpdf::Line;

use anyhow::{Context, Result};
use docx_rust::{
    document::{
        BodyContent, ParagraphContent, RunContent, Table, TableCellContent, TableRowContent,
    },
    DocxFile,
};
use env_logger;
use log::{debug, error, info};
use printpdf::*;
use std::{
    fs::File,
    io::{BufReader, BufWriter, Cursor, Read},
};

const PAGE_WIDTH: f32 = 210.0;
const PAGE_HEIGHT: f32 = 297.0;
const MARGIN: f32 = 10.0;
const LINE_HEIGHT: f32 = 6.0;
const PARAGRAPH_SPACING: f32 = 8.0;
const FONT_SIZE: f32 = 11.0;

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

#[derive(Debug)]
struct ImageContent {
    bytes: Vec<u8>,
}

#[derive(Debug)]
struct DocContent {
    text: String,
    image: Option<ImageContent>,
}

fn read_docx(docx_path: &str) -> Result<Vec<DocContent>> {
    debug!("Opening DOCX file: {}", docx_path);
    let doc = DocxFile::from_file(docx_path)
        .map_err(|e| anyhow::anyhow!("Failed to open DOCX file: {}: {:?}", docx_path, e))?;

    debug!("Parsing DOCX file");
    let docx = doc
        .parse()
        .map_err(|e| anyhow::anyhow!("Failed to parse DOCX file: {:?}", e))?;

    debug!("Processing DOCX content");
    let mut content_order = Vec::new();

    process_body_content(
        &docx.document.body.content,
        &docx,
        docx_path,
        &mut content_order,
    )?;

    debug!(
        "DOCX processing complete. Found {} content items",
        content_order.len()
    );
    Ok(content_order)
}

fn process_body_content(
    body_content: &Vec<BodyContent>,
    docx: &docx_rust::Docx,
    docx_path: &str,
    content_order: &mut Vec<DocContent>,
) -> Result<()> {
    for content in body_content {
        match content {
            BodyContent::Paragraph(paragraph) => {
                process_paragraph(paragraph, docx, docx_path, content_order)?;
            }
            BodyContent::Table(table) => {
                process_table(table, content_order)?;
            }
            _ => {}
        }
    }
    Ok(())
}

fn process_table(table: &Table, content_order: &mut Vec<DocContent>) -> Result<()> {
    let mut table_content = String::from("TABLE_START\n");

    for row in &table.rows {
        table_content.push('|');
        for cell in &row.cells {
            match cell {
                TableRowContent::TableCell(table_cell) => {
                    let mut cell_content = String::new();
                    for content in &table_cell.content {
                        match content {
                            TableCellContent::Paragraph(paragraph) => {
                                let mut paragraph_text = String::new();
                                process_paragraph_content(paragraph, &mut paragraph_text)?;
                                cell_content.push_str(&paragraph_text);
                            }
                        }
                    }
                    table_content.push_str(&cell_content);
                    table_content.push('|');
                }
                _ => {}
            }
        }
        table_content.push('\n');
    }

    table_content.push_str("TABLE_END\n");

    content_order.push(DocContent {
        text: table_content,
        image: None,
    });

    Ok(())
}

fn process_paragraph_content(
    paragraph: &docx_rust::document::Paragraph,
    paragraph_text: &mut String,
) -> Result<()> {
    for para_content in &paragraph.content {
        if let ParagraphContent::Run(run) = para_content {
            for run_content in &run.content {
                match run_content {
                    RunContent::Text(text) => {
                        paragraph_text.push_str(&text.text);
                    }
                    RunContent::Break(_) => {
                        paragraph_text.push(' ');
                    }
                    _ => {}
                }
            }
        }
    }
    Ok(())
}

fn process_paragraph(
    paragraph: &docx_rust::document::Paragraph,
    docx: &docx_rust::Docx,
    docx_path: &str,
    content_order: &mut Vec<DocContent>,
) -> Result<()> {
    let mut paragraph_text = String::new();
    for para_content in &paragraph.content {
        if let ParagraphContent::Run(run) = para_content {
            for run_content in &run.content {
                match run_content {
                    RunContent::Text(text) => {
                        paragraph_text.push_str(&text.text);
                    }
                    RunContent::Break(_) => {
                        paragraph_text.push('\n');
                    }
                    RunContent::Drawing(drawing) => {
                        if let Some(image_bytes) =
                            extract_image_from_drawing(drawing, docx, docx_path)?
                        {
                            content_order.push(DocContent {
                                text: String::new(),
                                image: Some(ImageContent { bytes: image_bytes }),
                            });
                        }
                    }
                    _ => {}
                }
            }
        }
    }
    if !paragraph_text.is_empty() {
        content_order.push(DocContent {
            text: paragraph_text,
            image: None,
        });
    }
    Ok(())
}

fn extract_image_from_drawing(
    drawing: &docx_rust::document::Drawing,
    docx: &docx_rust::Docx,
    docx_path: &str,
) -> Result<Option<Vec<u8>>> {
    if let Some(inline) = &drawing.inline {
        if let Some(graphic) = &inline.graphic {
            let rl_id = graphic.data.pic.fill.blip.embed.to_string();
            if let Some(relationships) = &docx.document_rels {
                if let Some(target) = relationships.get_target(&rl_id) {
                    return Ok(Some(extract_image_bytes(docx_path, &target)?));
                }
            }
        }
    }
    Ok(None)
}

fn extract_image_bytes(docx_path: &str, target: &str) -> Result<Vec<u8>> {
    let file = File::open(docx_path)
        .with_context(|| format!("Failed to open DOCX file: {}", docx_path))?;
    let mut zip = zip::ZipArchive::new(BufReader::new(file))
        .with_context(|| "Failed to create ZIP archive")?;

    let image_path = if target.starts_with("word/") {
        target.to_string()
    } else {
        format!("word/{}", target)
    };

    info!("Trying to open image file: {}", image_path);

    let mut image_file = zip
        .by_name(&image_path)
        .with_context(|| format!("Image not found in path: {}", image_path))?;

    let mut buffer = Vec::new();
    Read::read_to_end(&mut image_file, &mut buffer).with_context(|| "Failed to read image file")?;

    info!("Image file read successfully. Size: {} bytes", buffer.len());
    Ok(buffer)
}

fn convert_paragraphs_to_pdf(content: Vec<DocContent>, pdf_path: &str) -> Result<()> {
    debug!("Starting PDF conversion");
    let (doc, page1, layer1) = PdfDocument::new(
        "Converted Document",
        Mm(PAGE_WIDTH),
        Mm(PAGE_HEIGHT),
        "Layer 1",
    );
    let mut current_layer = doc.get_page(page1).get_layer(layer1);

    debug!("Adding built-in font");
    let font = doc.add_builtin_font(BuiltinFont::Helvetica)?;
    let font_bold = doc.add_builtin_font(BuiltinFont::HelveticaBold)?;

    let mut y_position = PAGE_HEIGHT - MARGIN;
    let max_width = PAGE_WIDTH - 2.0 * MARGIN;
    let indent = 2.0;

    debug!("Processing {} content items", content.len());
    for (index, item) in content.iter().enumerate() {
        if !item.text.is_empty() {
            if item.text.starts_with("TABLE_START") {
                y_position =
                    process_table_for_pdf(&item.text, &mut current_layer, y_position, &font)?;
            } else {
                let lines: Vec<&str> = item.text.split('\n').collect();
                for (line_index, line) in lines.iter().enumerate() {
                    let trimmed_line = line.trim();
                    if trimmed_line.is_empty() {
                        y_position -= PARAGRAPH_SPACING;
                        continue;
                    }

                    let (font_to_use, x_position) = if trimmed_line.starts_with('-') {
                        (&font, MARGIN + indent)
                    } else if line_index == 0 && lines.len() > 1 {
                        (&font_bold, MARGIN)
                    } else {
                        (&font, MARGIN)
                    };

                    let words: Vec<&str> = trimmed_line.split_whitespace().collect();
                    let mut current_line = String::new();
                    let mut current_width = 0.0;

                    for word in words {
                        let word_width = estimate_text_width(word, FONT_SIZE);
                        let space_width = estimate_text_width(" ", FONT_SIZE);

                        if current_width + word_width + space_width > max_width
                            && !current_line.is_empty()
                        {
                            debug!("Adding text at position {}", y_position);
                            current_layer.use_text(
                                current_line.clone(),
                                FONT_SIZE,
                                Mm(x_position),
                                Mm(y_position),
                                font_to_use,
                            );
                            y_position -= LINE_HEIGHT;
                            current_line.clear();
                            current_width = 0.0;
                        }

                        if !current_line.is_empty() {
                            current_line.push(' ');
                            current_width += space_width;
                        }
                        current_line.push_str(word);
                        current_width += word_width;
                    }

                    if !current_line.is_empty() {
                        debug!("Adding text at position {}", y_position);
                        current_layer.use_text(
                            current_line,
                            FONT_SIZE,
                            Mm(x_position),
                            Mm(y_position),
                            font_to_use,
                        );
                        y_position -= LINE_HEIGHT;
                    }
                }
                y_position -= PARAGRAPH_SPACING;
            }
        }

        if let Some(image) = &item.image {
            debug!("Processing image at index {}", index);

            let mut reader = Cursor::new(&image.bytes);

            let printpdf_image = match guess_format(&image.bytes)? {
                ImageFormat::Png => Image::try_from(PrintPdfPngDecoder::new(&mut reader)?)
                    .context("Falha ao converter a imagem PNG para o formato PDF")?,
                ImageFormat::Jpeg => Image::try_from(PrintPdfJpegDecoder::new(&mut reader)?)
                    .context("Falha ao converter a imagem JPEG para o formato PDF")?,
                _ => return Err(anyhow::anyhow!("Formato de imagem não suportado")),
            };

            let image_width = printpdf_image.image.width.into_pt(400.0);
            let image_height = printpdf_image.image.height.into_pt(400.0);

            let mut scale = (PAGE_WIDTH - 2.0 * MARGIN) / image_width.0;

            let max_height = y_position - MARGIN;
            if image_height.0 * scale > max_height {
                scale = max_height / image_height.0;
            }

            debug!("Escala da imagem: {}", scale);

            let scaled_width = image_width * scale;
            let scaled_height = image_height * scale;

            if y_position - scaled_height.0 < MARGIN {
                debug!("Adding new page for image");
                let (page, layer1) = doc.add_page(Mm(PAGE_WIDTH), Mm(PAGE_HEIGHT), "New Page");
                current_layer = doc.get_page(page).get_layer(layer1);
                y_position = PAGE_HEIGHT - MARGIN;
            }

            let x_position = (PAGE_WIDTH - scaled_width.0) / 2.0; // Centralizando a imagem

            printpdf_image.add_to_layer(
                current_layer.clone(),
                ImageTransform {
                    translate_x: Some(Mm(x_position)),
                    translate_y: Some(Mm(y_position - scaled_height.0)),
                    scale_x: Some(4.0),
                    scale_y: Some(4.0),
                    ..Default::default()
                },
            );

            y_position -= scaled_height.0 + PARAGRAPH_SPACING;
        }

        if y_position < MARGIN + 20.0 {
            debug!("Adding new page");
            let (page, layer1) = doc.add_page(Mm(PAGE_WIDTH), Mm(PAGE_HEIGHT), "New Page");
            current_layer = doc.get_page(page).get_layer(layer1);
            y_position = PAGE_HEIGHT - MARGIN;
        }
    }

    debug!("Saving PDF to {}", pdf_path);
    doc.save(&mut BufWriter::new(File::create(pdf_path)?))
        .with_context(|| format!("Failed to save PDF file: {}", pdf_path))?;

    let pdf_size = std::fs::metadata(pdf_path)?.len();
    info!("PDF saved successfully. File size: {} bytes", pdf_size);

    Ok(())
}

fn estimate_text_width(text: &str, font_size: f32) -> f32 {
    let average_char_width = font_size * 0.25;
    text.len() as f32 * average_char_width
}

fn convert_docx_to_pdf(docx_path: &str, pdf_path: &str) -> Result<()> {
    debug!("Reading DOCX file: {}", docx_path);
    let content =
        read_docx(docx_path).with_context(|| format!("Failed to read DOCX file: {}", docx_path))?;

    info!("Successfully read DOCX file. Converting to PDF...");

    convert_paragraphs_to_pdf(content, pdf_path)
        .with_context(|| format!("Failed to convert paragraphs to PDF: {}", pdf_path))?;

    Ok(())
}
fn process_table_for_pdf(
    table_content: &str,
    current_layer: &mut PdfLayerReference,
    mut y_position: f32,
    font: &IndirectFontRef,
) -> Result<f32> {
    let rows: Vec<&str> = table_content.split('\n').collect();
    let num_columns = rows[1].split('|').count() - 2;
    let column_width = (PAGE_WIDTH - 2.0 * MARGIN) / num_columns as f32;
    let initial_y = y_position;

    draw_horizontal_line(current_layer, MARGIN, y_position, num_columns, column_width);

    for (_, row) in rows.iter().enumerate().skip(1) {
        if row.trim() == "TABLE_END" {
            break;
        }

        y_position -= LINE_HEIGHT;

        let cells: Vec<&str> = row.split('|').collect();
        for (col_index, cell) in cells.iter().enumerate().skip(1).take(num_columns) {
            let x = MARGIN + (col_index - 1) as f32 * column_width;
            current_layer.use_text(
                cell.trim().to_string(),
                FONT_SIZE,
                Mm(x + 13.0),
                Mm(y_position + 2.0),
                font,
            );

            draw_vertical_line(current_layer, x, initial_y, y_position);
        }
        draw_horizontal_line(current_layer, MARGIN, y_position, num_columns, column_width);
    }

    draw_vertical_line(
        current_layer,
        MARGIN + num_columns as f32 * column_width,
        initial_y,
        y_position,
    );

    draw_horizontal_line(current_layer, MARGIN, y_position, num_columns, column_width);

    Ok(y_position)
}

fn draw_horizontal_line(
    layer: &mut PdfLayerReference,
    x: f32,
    y: f32,
    num_columns: usize,
    column_width: f32,
) {
    let line = Line {
        points: vec![
            (Point::new(Mm(x), Mm(y)), false),
            (
                Point::new(Mm(x + num_columns as f32 * column_width), Mm(y)),
                false,
            ),
        ],
        is_closed: false,
    };
    layer.add_line(line);
}

fn draw_vertical_line(layer: &mut PdfLayerReference, x: f32, y_start: f32, y_end: f32) {
    let line = Line {
        points: vec![
            (Point::new(Mm(x), Mm(y_start)), false),
            (Point::new(Mm(x), Mm(y_end)), false),
        ],
        is_closed: false,
    };
    layer.add_line(line);
}
