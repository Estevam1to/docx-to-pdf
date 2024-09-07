use crate::utils::{DocContent, ImageContent};

use anyhow::{Context, Result};
use docx_rust::{
    document::{
        BodyContent, ParagraphContent, RunContent, Table, TableCellContent, TableRowContent,
    },
    DocxFile,
};
use log::{debug, info};
use std::{
    fs::File,
    io::{BufReader, Read},
};


pub fn read_docx(docx_path: &str) -> Result<Vec<DocContent>> {
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