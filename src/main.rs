use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::Path;
use zip::read::ZipArchive;
use docx_rs::{
    Docx,
    Document,
    Run,
    RunProperty,
    read_docx
};
use printpdf::*;
use log::{info, error};
use env_logger::Env;
use ::image::DynamicImage;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum ConversionError{
    #[error("IO error: {0}")]
    Io(#[from] std::io::Error),
    #[error("Zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("Docx parsing error: {0}")]
    Docx(#[from] docx_rs::DocxError),
    #[error("Image processing error: {0}")]
    Image(#[from] image::DynamicImage),
    #[error("PDF creation error: {0}")]
    Pdf(String),
    #[error("Invalid input file: {0}")]
    InvalidInput(String),
}

struct Config{
    input_path: String,
    output_path: String,
    page_width: f32,
    page_height: f32,
    margin: f32,
}

impl Config{
    fn new(input_path: &str, output_path: &str) -> Self{
        Config{
            input_path: input_path.to_string(),
            output_path: output_path.to_string(),
            page_width: 210.0,
            page_height: 297.0,
            margin: 20.0,
        }
    }
}

fn main() -> Result<(), ConversionError> {
    //Initializing logger
    env_logger::Builder::from_env(Env::default().default_filter_or("info")).init();
        
        //Parse command-line arguments
        let args: Vec<String> = std::env::args().collect();
        if args.len() != 3 {
            eprintln!("Usage: {} <input.docx> <output.pdf>", args[0]);
            std::process::exit(1);
        }

        let config = Config::new(&args[1], &args[2]);

        //This validates the input file
        if !Path::new(&config.input_path).exists() || !config.input_path.ends_with(".docx") {
            return Err(ConversionError::InvalidInput("Error: Invalid input file".to_string()));
        }

        info!("Starting conversion from {} to {}", config.input_path, config.output_path);

        //Reads and parse .docx file
        let docx_content = fs::read(&config.input_path)?;
        let docx = read_docx(&docx_content)?;

        //Extracts images
        
        let image = extract_images(&config.input_path) ?;

        //Generate PDF
        create_pdf(&docx, &image, &config)?;

        info!("Conversion completed successfully.", config.output_path);
        Ok(())
}

fn extract_images(docx_path:&str) -> Result<Vec<(String, DynamicImage)>, ConversionError>{
    let file = File::open(docx_path)?;
    let mut archive = ZipArchive::new(file)?;
    let mut images = Vec::new();

    for i in 0..archive.len(){
        let mut zip_file = archive.by_index(i)  ?;
        let file_name = zip_file.name().to_string();
        if file_name.starts_with("word/media"){
            let mut buffer = Vec::new();
            zip_file.read_to_end(&mut buffer)?;
            if let Ok(img) = image::load_from_memory(&buffer) {
                images.push((file_name, img));
                info!("Extracted image: {}", file_name);
        }
    }
}
    Ok(images)
}

fn create_pdf (docx: &Docx, images:&[(String, DynamicImage)], config: &Config) -> Result<(), CoversionError>{
    let (doc, page1, layer1) = PdfDocument::new(
        "Word to PDF",
        Mm(config.page_width),
        Mm(config.page_height), 
        "Layer 1",

    );
    let mut current_layer  = doc.get_page(page1).get_layer(layer1);

    //Load fonts
    let regular_font = doc.add_builtin_font(BuiltinFont::Helvetica)?;
    let bold_font = doc.add_builtin_font(builtin_font::HelveticaBold)?;
    let italic_font = doc.add_builtin_font(BuiltinFont::HelveticaOblique)?;

    let mut y_position = config.page_height - config.margin;
    let line_height = 12.0;
    let font_size = 12.0;

    //Processes document content
    let Document { children, .. } = &docx.document;
    for child in children {
        match child {
            docx_rs::DocumentChild::Paragraph(paragraph) => {
                for run in &paragraph.runs {
                    if let Run::Text { content, properties } = run {
                        let font = match properties {
                            RunProperty::Bold => bold_font,
                            RunProperty::Italic => italic_font,
                            _ => regular_font,
                        };
                        //Split text into lines that's if needed
                        let words = content.split_whitespace();
                        let mut current_line = String::new();
                        for word in words {
                            if current_line.len() + word.len() < 80 {
                                current_line.push_str(word);
                                current_line.push(' ');
                            } else {
                                current_layer.use_text(
                                    &current_line,
                                    font_size,
                                    Mm(config.margin),
                                    Mm(config.page_height - y_position),
                                    &font,
                                );
                                y_position -= line_height;
                                current_line = format!("{} ", word);

                                //Checks if the data has a page break
                                if y_position < config.margin {
                                    let (new_page, new_layer) = doc.add_page(
                                        Mm(config.page_width),
                                        Mm(config.page_height),
                                        "Layer 1",
                                    );
                                    current_layer = doc.get_page(new_page).get_layer(new_layer);
                                    y_position = config.page_height - config.margin;
                                }
                            }
                        }
                        if !current_line.is_empty() {
                            current_layer.use_text(
                                &current_line,
                                font_size,
                                Mm(config.margin),
                                Mm(y_position),
                                &font,
                            );
                            y_position -= line_height;
                        }
                    }
                }
                y_position -= line_height;
            }
                _ => {}
            }
            
        }
    

    // Adds the images if they exist
    for (name, img) in images{
        if y_position < config.margin + 50+0{
            let (new_page, new_layer) = doc.add_page(
                Mm(config.page_width),
                Mm(config.page_height),
                "Layer 1",
            );
            current_layer = doc.get_page(new_page).get_layer(new_layer);
            y_position = config.page_height - config.margin;
        }

        let (width, height) = img.dimensions();
        let scale = (config.page_width - 2.0 * config.margin) / width as f32;
        // Convert the DynamicImage to RGB8 and get raw bytes
        let rgb_image = img.to_rgb8();
        let (img_width, img_height) = rgb_image.dimensions();
        let image_bytes = rgb_image.into_raw();

        // Create an Image in the PDF
        let image = Image::from_rgb(
            img_width as usize,
            img_height as usize,
            &image_bytes,
        );

        // Calculate scaled width and height
        let scaled_width = (img_width as f32) * scale;
        let scaled_height = (img_height as f32) * scale;

        // Add the image to the current layer
        image.add_to_layer(
            current_layer.clone(),
            ImageTransform {
                translate_x: Some(Mm(config.margin)),
                translate_y: Some(Mm(y_position - scaled_height)),
                rotate: None,
                scale_x: Some(scale as f32),
                scale_y: Some(scale as f32),
                dpi: None,
            },
        );
        y_position -= scaled_height + 10.0;
    }
    //Saves the PDF

    let mut file = File::create(&config.output_path)?;
    file.write_all(&doc.save_to_bytes()?)?;
    Ok(())
}