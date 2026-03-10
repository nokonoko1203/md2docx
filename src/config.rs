use anyhow::Result;
use serde::Deserialize;
use std::path::Path;

#[derive(Debug, Deserialize)]
pub struct BulletConfig {
    #[serde(default = "default_bullet_level0")]
    pub level0: String,
    #[serde(default = "default_bullet_level1")]
    pub level1: String,
    #[serde(default = "default_bullet_level2")]
    pub level2: String,
}

#[derive(Debug, Deserialize, Default)]
pub struct Config {
    #[serde(default)]
    pub fonts: FontConfig,
    #[serde(default)]
    pub sizes: SizeConfig,
    #[serde(default)]
    pub page: PageConfig,
    #[serde(default)]
    pub indent: IndentConfig,
    #[serde(default)]
    pub bullet: BulletConfig,
    #[serde(default)]
    pub numbering: NumberingConfig,
}

#[derive(Debug, Deserialize)]
pub struct FontConfig {
    #[serde(default = "default_body_ja")]
    pub body_ja: String,
    #[serde(default = "default_body_en")]
    pub body_en: String,
    #[serde(default = "default_heading_ja")]
    pub heading_ja: String,
    #[serde(default = "default_heading_en")]
    pub heading_en: String,
}

#[derive(Debug, Deserialize)]
pub struct SizeConfig {
    #[serde(default = "default_body_size")]
    pub body: f64,
    #[serde(default = "default_h1_size")]
    pub heading1: f64,
    #[serde(default = "default_h2_size")]
    pub heading2: f64,
    #[serde(default = "default_h3_size")]
    pub heading3: f64,
    #[serde(default = "default_h4_size")]
    pub heading4: f64,
    #[serde(default = "default_h5_size")]
    pub heading5: f64,
}

#[derive(Debug, Deserialize)]
pub struct PageConfig {
    #[serde(default = "default_page_width")]
    pub width: u32,
    #[serde(default = "default_page_height")]
    pub height: u32,
    #[serde(default = "default_page_margin_top")]
    pub margin_top: i32,
    #[serde(default = "default_page_margin_right")]
    pub margin_right: i32,
    #[serde(default = "default_page_margin_bottom")]
    pub margin_bottom: i32,
    #[serde(default = "default_page_margin_left")]
    pub margin_left: i32,
    #[serde(default = "default_page_margin_header")]
    pub margin_header: i32,
    #[serde(default = "default_page_margin_footer")]
    pub margin_footer: i32,
    #[serde(default = "default_page_margin_gutter")]
    pub margin_gutter: i32,
}

#[derive(Debug, Deserialize)]
pub struct IndentConfig {
    #[serde(default = "default_indent_body_left")]
    pub body_left: i32,
    #[serde(default = "default_indent_body_first_line")]
    pub body_first_line: i32,
    #[serde(default = "default_indent_body_right")]
    pub body_right: i32,
    #[serde(default = "default_indent_body_left_chars")]
    pub body_left_chars: i32,
    #[serde(default = "default_indent_heading4_left")]
    pub heading4_left: i32,
    #[serde(default = "default_indent_heading4_hanging")]
    pub heading4_hanging: i32,
}

fn default_body_ja() -> String {
    "游明朝".to_string()
}
fn default_body_en() -> String {
    "Century".to_string()
}
fn default_heading_ja() -> String {
    "游ゴシック".to_string()
}
fn default_heading_en() -> String {
    "Century".to_string()
}
fn default_body_size() -> f64 {
    10.5
}
fn default_h1_size() -> f64 {
    14.0
}
fn default_h2_size() -> f64 {
    12.0
}
fn default_h3_size() -> f64 {
    11.0
}
fn default_h4_size() -> f64 {
    11.0
}
fn default_h5_size() -> f64 {
    10.5
}

fn default_page_width() -> u32 {
    11_906
}
fn default_page_height() -> u32 {
    16_838
}
fn default_page_margin_top() -> i32 {
    1_985
}
fn default_page_margin_right() -> i32 {
    1_701
}
fn default_page_margin_bottom() -> i32 {
    1_701
}
fn default_page_margin_left() -> i32 {
    1_701
}
fn default_page_margin_header() -> i32 {
    851
}
fn default_page_margin_footer() -> i32 {
    992
}
fn default_page_margin_gutter() -> i32 {
    0
}

fn default_indent_body_left() -> i32 {
    210
}
fn default_indent_body_first_line() -> i32 {
    210
}
fn default_indent_body_right() -> i32 {
    210
}
fn default_indent_body_left_chars() -> i32 {
    100
}
fn default_indent_heading4_left() -> i32 {
    709
}
fn default_indent_heading4_hanging() -> i32 {
    709
}

#[derive(Debug, Deserialize)]
pub struct NumberingConfig {
    #[serde(default = "default_figure_format")]
    pub figure_format: String,
    #[serde(default = "default_table_format")]
    pub table_format: String,
}

fn default_figure_format() -> String {
    "sequential".to_string()
}
fn default_table_format() -> String {
    "sequential".to_string()
}

impl Default for NumberingConfig {
    fn default() -> Self {
        Self {
            figure_format: default_figure_format(),
            table_format: default_table_format(),
        }
    }
}

fn default_bullet_level0() -> String {
    "●".to_string()
}
fn default_bullet_level1() -> String {
    "■".to_string()
}
fn default_bullet_level2() -> String {
    "▲".to_string()
}

impl Default for FontConfig {
    fn default() -> Self {
        Self {
            body_ja: default_body_ja(),
            body_en: default_body_en(),
            heading_ja: default_heading_ja(),
            heading_en: default_heading_en(),
        }
    }
}

impl Default for SizeConfig {
    fn default() -> Self {
        Self {
            body: default_body_size(),
            heading1: default_h1_size(),
            heading2: default_h2_size(),
            heading3: default_h3_size(),
            heading4: default_h4_size(),
            heading5: default_h5_size(),
        }
    }
}

impl Default for PageConfig {
    fn default() -> Self {
        Self {
            width: default_page_width(),
            height: default_page_height(),
            margin_top: default_page_margin_top(),
            margin_right: default_page_margin_right(),
            margin_bottom: default_page_margin_bottom(),
            margin_left: default_page_margin_left(),
            margin_header: default_page_margin_header(),
            margin_footer: default_page_margin_footer(),
            margin_gutter: default_page_margin_gutter(),
        }
    }
}

impl Default for IndentConfig {
    fn default() -> Self {
        Self {
            body_left: default_indent_body_left(),
            body_first_line: default_indent_body_first_line(),
            body_right: default_indent_body_right(),
            body_left_chars: default_indent_body_left_chars(),
            heading4_left: default_indent_heading4_left(),
            heading4_hanging: default_indent_heading4_hanging(),
        }
    }
}

impl Default for BulletConfig {
    fn default() -> Self {
        Self {
            level0: default_bullet_level0(),
            level1: default_bullet_level1(),
            level2: default_bullet_level2(),
        }
    }
}

impl Config {
    pub fn load(path: &Path) -> Result<Self> {
        let content = std::fs::read_to_string(path)?;
        let config: Config = toml::from_str(&content)?;
        Ok(config)
    }
}
