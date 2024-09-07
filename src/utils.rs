#[derive(Debug)]
pub struct ImageContent {
    pub bytes: Vec<u8>,
}

#[derive(Debug)]
pub struct DocContent {
    pub text: String,
    pub image: Option<ImageContent>,
}

pub fn estimate_text_width(text: &str, font_size: f32) -> f32 {
    let average_char_width = font_size * 0.25;
    text.len() as f32 * average_char_width
}

