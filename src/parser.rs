use anyhow::{Result, bail};
use pulldown_cmark::{Event, HeadingLevel, Options, Parser, Tag, TagEnd};

use crate::ir::{Alignment, Block, Inline, ListItem};

const PAGE_BREAK_DIRECTIVE: &str = r"\pagebreak";

pub fn parse_markdown(input: &str) -> Result<Vec<Block>> {
    let mut options = Options::empty();
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TASKLISTS);

    let parser = Parser::new_ext(input, options);
    let events: Vec<Event> = parser.collect();

    let converter = EventConverter::new();
    let blocks = converter.convert(&events);
    validate_page_break_usage(&blocks)?;
    Ok(blocks)
}

struct EventConverter {
    blocks: Vec<Block>,
    inline_stack: Vec<Vec<Inline>>,
    list_stack: Vec<ListContext>,
    table_state: Option<TableState>,
    in_block_quote: bool,
    block_quote_blocks: Vec<Block>,
    current_image_path: Option<String>,
    current_code_lang: Option<String>,
    current_link_url: Option<String>,
}

struct ListContext {
    ordered: bool,
    start: u64,
    items: Vec<ListItem>,
    current_item_inlines: Vec<Inline>,
    current_item_children: Vec<Block>,
}

struct TableState {
    headers: Vec<Vec<Inline>>,
    rows: Vec<Vec<Vec<Inline>>>,
    current_row: Vec<Vec<Inline>>,
    current_cell: Vec<Inline>,
    in_header: bool,
    alignments: Vec<Alignment>,
}

impl EventConverter {
    fn new() -> Self {
        Self {
            blocks: Vec::new(),
            inline_stack: Vec::new(),
            list_stack: Vec::new(),
            table_state: None,
            in_block_quote: false,
            block_quote_blocks: Vec::new(),
            current_image_path: None,
            current_code_lang: None,
            current_link_url: None,
        }
    }

    fn convert(mut self, events: &[Event]) -> Vec<Block> {
        for event in events {
            self.process_event(event);
        }
        self.blocks
    }

    fn push_inline(&mut self, inline: Inline) {
        if let Some(stack) = self.inline_stack.last_mut() {
            stack.push(inline);
        } else if let Some(ref mut table) = self.table_state {
            table.current_cell.push(inline);
        } else if let Some(list_ctx) = self.list_stack.last_mut() {
            list_ctx.current_item_inlines.push(inline);
        }
    }

    fn process_event(&mut self, event: &Event) {
        match event {
            Event::Start(tag) => self.handle_start(tag),
            Event::End(tag) => self.handle_end(tag),
            Event::Text(text) => {
                self.push_inline(Inline::Text(text.to_string()));
            }
            Event::Code(code) => {
                self.push_inline(Inline::Code(code.to_string()));
            }
            Event::SoftBreak => {
                self.push_inline(Inline::SoftBreak);
            }
            Event::HardBreak => {
                self.push_inline(Inline::HardBreak);
            }
            Event::Rule => {
                self.add_block(Block::ThematicBreak);
            }
            _ => {}
        }
    }

    fn handle_start(&mut self, tag: &Tag) {
        match tag {
            Tag::Heading { level, .. } => {
                self.inline_stack.push(Vec::new());
                let _ = level; // level is used in handle_end
            }
            Tag::Paragraph => {
                self.inline_stack.push(Vec::new());
            }
            Tag::Emphasis => {
                self.inline_stack.push(Vec::new());
            }
            Tag::Strong => {
                self.inline_stack.push(Vec::new());
            }
            Tag::Link { dest_url, .. } => {
                self.inline_stack.push(Vec::new());
                self.current_link_url = Some(dest_url.to_string());
            }
            Tag::List(start) => {
                let ordered = start.is_some();
                let start_num = start.unwrap_or(1);
                self.list_stack.push(ListContext {
                    ordered,
                    start: start_num,
                    items: Vec::new(),
                    current_item_inlines: Vec::new(),
                    current_item_children: Vec::new(),
                });
            }
            Tag::Item => {
                if let Some(list_ctx) = self.list_stack.last_mut() {
                    list_ctx.current_item_inlines = Vec::new();
                    list_ctx.current_item_children = Vec::new();
                }
            }
            Tag::Table(alignments) => {
                let aligns = alignments
                    .iter()
                    .map(|a| match a {
                        pulldown_cmark::Alignment::Left => Alignment::Left,
                        pulldown_cmark::Alignment::Center => Alignment::Center,
                        pulldown_cmark::Alignment::Right => Alignment::Right,
                        pulldown_cmark::Alignment::None => Alignment::None,
                    })
                    .collect();
                self.table_state = Some(TableState {
                    headers: Vec::new(),
                    rows: Vec::new(),
                    current_row: Vec::new(),
                    current_cell: Vec::new(),
                    in_header: false,
                    alignments: aligns,
                });
            }
            Tag::TableHead => {
                if let Some(ref mut state) = self.table_state {
                    state.in_header = true;
                    state.current_row = Vec::new();
                }
            }
            Tag::TableRow => {
                if let Some(ref mut state) = self.table_state {
                    state.current_row = Vec::new();
                }
            }
            Tag::TableCell => {
                if let Some(ref mut state) = self.table_state {
                    state.current_cell = Vec::new();
                }
            }
            Tag::BlockQuote(_) => {
                self.in_block_quote = true;
                self.block_quote_blocks = Vec::new();
            }
            Tag::CodeBlock(kind) => {
                self.current_code_lang = match kind {
                    pulldown_cmark::CodeBlockKind::Fenced(lang) => {
                        let l = lang.to_string();
                        if l.is_empty() { None } else { Some(l) }
                    }
                    pulldown_cmark::CodeBlockKind::Indented => None,
                };
                self.inline_stack.push(Vec::new());
            }
            Tag::Image { dest_url, .. } => {
                self.inline_stack.push(Vec::new());
                self.current_image_path = Some(dest_url.to_string());
            }
            _ => {}
        }
    }

    fn handle_end(&mut self, tag: &TagEnd) {
        match tag {
            TagEnd::Heading(level) => {
                let content = self.inline_stack.pop().unwrap_or_default();
                let lvl = heading_level_to_u8(level);
                self.add_block(Block::Heading {
                    level: lvl,
                    content,
                });
            }
            TagEnd::Paragraph => {
                let content = self.inline_stack.pop().unwrap_or_default();
                if is_page_break_paragraph(&content) {
                    self.add_block(Block::PageBreak);
                } else if !content.is_empty() {
                    self.add_block(Block::Paragraph { content });
                }
            }
            TagEnd::Emphasis => {
                let children = self.inline_stack.pop().unwrap_or_default();
                self.push_inline(Inline::Italic(children));
            }
            TagEnd::Strong => {
                let children = self.inline_stack.pop().unwrap_or_default();
                self.push_inline(Inline::Bold(children));
            }
            TagEnd::Link => {
                let text = self.inline_stack.pop().unwrap_or_default();
                let url = self.current_link_url.take().unwrap_or_default();
                if url.is_empty() {
                    for inline in text {
                        self.push_inline(inline);
                    }
                } else {
                    self.push_inline(Inline::Link { text, url });
                }
            }
            TagEnd::List(_ordered) => {
                if let Some(list_ctx) = self.list_stack.pop() {
                    let block = if list_ctx.ordered {
                        Block::OrderedList {
                            items: list_ctx.items,
                            start: list_ctx.start,
                        }
                    } else {
                        Block::BulletList {
                            items: list_ctx.items,
                        }
                    };
                    self.add_block(block);
                }
            }
            TagEnd::Item => {
                if let Some(list_ctx) = self.list_stack.last_mut() {
                    let item = ListItem {
                        content: std::mem::take(&mut list_ctx.current_item_inlines),
                        children: std::mem::take(&mut list_ctx.current_item_children),
                    };
                    list_ctx.items.push(item);
                }
            }
            TagEnd::Table => {
                if let Some(state) = self.table_state.take() {
                    self.add_block(Block::Table {
                        headers: state.headers,
                        rows: state.rows,
                        alignments: state.alignments,
                    });
                }
            }
            TagEnd::TableHead => {
                if let Some(ref mut state) = self.table_state {
                    state.headers = std::mem::take(&mut state.current_row);
                    state.in_header = false;
                }
            }
            TagEnd::TableRow => {
                if let Some(ref mut state) = self.table_state
                    && !state.in_header
                {
                    let row = std::mem::take(&mut state.current_row);
                    state.rows.push(row);
                }
            }
            TagEnd::TableCell => {
                if let Some(ref mut state) = self.table_state {
                    let cell = std::mem::take(&mut state.current_cell);
                    state.current_row.push(cell);
                }
            }
            TagEnd::BlockQuote(_) => {
                let children = std::mem::take(&mut self.block_quote_blocks);
                self.in_block_quote = false;
                self.add_block(Block::BlockQuote { children });
            }
            TagEnd::CodeBlock => {
                let content = self.inline_stack.pop().unwrap_or_default();
                let code: String = content
                    .iter()
                    .map(|i| match i {
                        Inline::Text(s) => s.as_str(),
                        _ => "",
                    })
                    .collect();
                let lang = self.current_code_lang.take();
                self.add_block(Block::CodeBlock { lang, code });
            }
            TagEnd::Image => {
                let alt_inlines = self.inline_stack.pop().unwrap_or_default();
                let alt: String = alt_inlines.iter().map(|i| i.to_plain_text()).collect();
                let path = self.current_image_path.take().unwrap_or_default();
                self.add_block(Block::Image { alt, path });
            }
            _ => {}
        }
    }

    fn add_block(&mut self, block: Block) {
        if self.in_block_quote {
            self.block_quote_blocks.push(block);
        } else if !self.list_stack.is_empty() {
            // リスト内のネストされたブロック
            if let Some(list_ctx) = self.list_stack.last_mut() {
                // Paragraph内のインラインをリストアイテムに移動
                match &block {
                    Block::Paragraph { content } => {
                        if list_ctx.current_item_inlines.is_empty() {
                            list_ctx.current_item_inlines = content.clone();
                        } else {
                            list_ctx.current_item_children.push(block);
                        }
                    }
                    _ => {
                        list_ctx.current_item_children.push(block);
                    }
                }
            }
        } else {
            self.blocks.push(block);
        }
    }
}

fn is_page_break_paragraph(content: &[Inline]) -> bool {
    matches!(content, [Inline::Text(text)] if text.trim() == PAGE_BREAK_DIRECTIVE)
}

fn validate_page_break_usage(blocks: &[Block]) -> Result<()> {
    for block in blocks {
        validate_block(block)?;
    }
    Ok(())
}

fn validate_block(block: &Block) -> Result<()> {
    match block {
        Block::Heading { content, .. } | Block::Paragraph { content } => {
            validate_inlines(content)?;
        }
        Block::BulletList { items } | Block::OrderedList { items, .. } => {
            for item in items {
                validate_inlines(&item.content)?;
                for child in &item.children {
                    validate_block(child)?;
                }
            }
        }
        Block::Table { headers, rows, .. } => {
            for cell in headers.iter().flatten() {
                validate_inline(cell)?;
            }
            for row in rows {
                for cell in row {
                    validate_inlines(cell)?;
                }
            }
        }
        Block::CodeBlock { code, .. } => {
            ensure_no_page_break_directive(code)?;
        }
        Block::Image { alt, .. } => {
            ensure_no_page_break_directive(alt)?;
        }
        Block::BlockQuote { children } => {
            for child in children {
                validate_block(child)?;
            }
        }
        Block::PageBreak | Block::ThematicBreak => {}
    }
    Ok(())
}

fn validate_inlines(inlines: &[Inline]) -> Result<()> {
    for inline in inlines {
        validate_inline(inline)?;
    }
    Ok(())
}

fn validate_inline(inline: &Inline) -> Result<()> {
    match inline {
        Inline::Text(text) | Inline::Code(text) => ensure_no_page_break_directive(text)?,
        Inline::Bold(children) | Inline::Italic(children) => validate_inlines(children)?,
        Inline::Link { text, url } => {
            validate_inlines(text)?;
            ensure_no_page_break_directive(url)?;
        }
        Inline::SoftBreak | Inline::HardBreak => {}
    }
    Ok(())
}

fn ensure_no_page_break_directive(text: &str) -> Result<()> {
    if text.contains(PAGE_BREAK_DIRECTIVE) {
        bail!(
            "`{}`は段落単位で単独指定した場合のみ利用できます",
            PAGE_BREAK_DIRECTIVE
        );
    }
    Ok(())
}

fn heading_level_to_u8(level: &HeadingLevel) -> u8 {
    match level {
        HeadingLevel::H1 => 1,
        HeadingLevel::H2 => 2,
        HeadingLevel::H3 => 3,
        HeadingLevel::H4 => 4,
        HeadingLevel::H5 => 5,
        HeadingLevel::H6 => 6,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn preserves_link_url_in_inline_ir() {
        let blocks = parse_markdown("[Rust](https://www.rust-lang.org/)").unwrap();
        assert_eq!(blocks.len(), 1);

        match &blocks[0] {
            Block::Paragraph { content } => match &content[0] {
                Inline::Link { text, url } => {
                    assert_eq!(url, "https://www.rust-lang.org/");
                    assert_eq!(text.len(), 1);
                    assert!(matches!(&text[0], Inline::Text(t) if t == "Rust"));
                }
                other => panic!("unexpected inline: {other:?}"),
            },
            other => panic!("unexpected block: {other:?}"),
        }
    }

    #[test]
    fn does_not_mix_urls_between_multiple_links() {
        let blocks = parse_markdown("[A](https://a.example) [B](https://b.example)").unwrap();
        assert_eq!(blocks.len(), 1);

        match &blocks[0] {
            Block::Paragraph { content } => {
                let links: Vec<&Inline> = content
                    .iter()
                    .filter(|i| matches!(i, Inline::Link { .. }))
                    .collect();
                assert_eq!(links.len(), 2);

                match links[0] {
                    Inline::Link { url, .. } => assert_eq!(url, "https://a.example"),
                    _ => unreachable!(),
                }
                match links[1] {
                    Inline::Link { url, .. } => assert_eq!(url, "https://b.example"),
                    _ => unreachable!(),
                }
            }
            other => panic!("unexpected block: {other:?}"),
        }
    }

    #[test]
    fn parses_page_break_directive_as_dedicated_block() {
        let blocks = parse_markdown("\\pagebreak").unwrap();
        assert_eq!(blocks.len(), 1);
        assert!(matches!(blocks[0], Block::PageBreak));
    }

    #[test]
    fn rejects_page_break_directive_inside_regular_paragraph() {
        let error = parse_markdown("before \\pagebreak after").unwrap_err();
        assert!(
            error
                .to_string()
                .contains(r"`\pagebreak`は段落単位で単独指定した場合のみ利用できます")
        );
    }
}
