use anyhow::{Context, Result};
use log::{debug, info};
use printpdf::image_crate::codecs::jpeg::JpegDecoder as PrintPdfJpegDecoder;
use printpdf::image_crate::codecs::png::PngDecoder as PrintPdfPngDecoder;
use printpdf::image_crate::{guess_format, ImageFormat};
use printpdf::*;
use std::io::Cursor;
use std::{fs::File, io::BufWriter};

use crate::utils::{estimate_text_width, DocContent};
use crate::{FONT_SIZE, LINE_HEIGHT, MARGIN, PAGE_HEIGHT, PAGE_WIDTH, PARAGRAPH_SPACING};

pub fn convert_paragraphs_to_pdf(content: Vec<DocContent>, pdf_path: &str) -> Result<()> {
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
