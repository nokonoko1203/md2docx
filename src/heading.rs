use crate::ir::Inline;

/// 見出し番号管理
/// H1 → X, H2 → X.X, H3 → X.X.X, H4 → (N), H5 → ①
pub struct HeadingManager {
    counters: [u32; 5], // h1..h5
}

impl HeadingManager {
    pub fn new() -> Self {
        Self { counters: [0; 5] }
    }

    /// 見出しレベルに応じて番号を更新し、フォーマットされた番号文字列を返す。
    /// 見出しテキストに既に番号が含まれている場合はそれを尊重する。
    pub fn next_heading(&mut self, level: u8, content: &[Inline]) -> String {
        let plain_text: String = content.iter().map(|i| i.to_plain_text()).collect();

        // 既存の番号を検出する
        if let Some(existing) = self.detect_existing_number(level, &plain_text) {
            self.sync_counters(level, &existing);
            return existing;
        }

        // 自動採番
        self.increment(level);
        self.format_number(level)
    }

    /// テキストから見出し番号部分を除去し、タイトルのみを返す。
    /// `detect_existing_number` を内部で使い、番号フォーマット判定ロジックを一元化する。
    pub fn strip_number(&self, level: u8, text: &str) -> String {
        let trimmed = text.trim();
        if self.detect_existing_number(level, trimmed).is_some() {
            match level {
                1 => {
                    // "8 タイトル" → "タイトル"
                    if let Some(rest) = trimmed.split_once(' ') {
                        return rest.1.to_string();
                    }
                }
                2 => {
                    // "8.1 タイトル" → "タイトル"
                    if let Some(rest) = trimmed.split_once(' ') {
                        return rest.1.to_string();
                    }
                }
                3 => {
                    // "8.1.1 タイトル" → "タイトル"
                    if let Some(rest) = trimmed.split_once(' ') {
                        return rest.1.to_string();
                    }
                }
                4 => {
                    // "(1) タイトル" → "タイトル"
                    if let Some(end) = trimmed.find(')') {
                        return trimmed[end + 1..].trim_start().to_string();
                    }
                }
                5 => {
                    // "① タイトル" → "タイトル"
                    let mut chars = trimmed.chars();
                    chars.next(); // 丸数字をスキップ
                    return chars.as_str().trim_start().to_string();
                }
                _ => {}
            }
        }
        trimmed.to_string()
    }

    /// 現在の H1 カウンタ値を返す。
    pub fn current_h1_number(&self) -> u32 {
        self.counters[0]
    }

    fn increment(&mut self, level: u8) {
        let idx = (level as usize).saturating_sub(1).min(4);
        self.counters[idx] += 1;
        // 下位レベルをリセット
        for i in (idx + 1)..5 {
            self.counters[i] = 0;
        }
    }

    fn format_number(&self, level: u8) -> String {
        match level {
            1 => format!("{}", self.counters[0]),
            2 => format!("{}.{}", self.counters[0], self.counters[1]),
            3 => format!(
                "{}.{}.{}",
                self.counters[0], self.counters[1], self.counters[2]
            ),
            4 => {
                let n = self.counters[3];
                format!("({})", n)
            }
            5 => {
                let n = self.counters[4];
                num_to_circled(n)
            }
            _ => String::new(),
        }
    }

    fn detect_existing_number(&self, level: u8, text: &str) -> Option<String> {
        let trimmed = text.trim();
        match level {
            1 => {
                // "8 タイトル" → "8"
                if let Some(num_str) = trimmed.split_whitespace().next() {
                    if num_str.parse::<u32>().is_ok() {
                        return Some(num_str.to_string());
                    }
                }
                None
            }
            2 => {
                // "8.1 タイトル" → "8.1"
                if let Some(num_str) = trimmed.split_whitespace().next() {
                    let parts: Vec<&str> = num_str.split('.').collect();
                    if parts.len() == 2
                        && parts[0].parse::<u32>().is_ok()
                        && parts[1].parse::<u32>().is_ok()
                    {
                        return Some(num_str.to_string());
                    }
                }
                None
            }
            3 => {
                // "8.1.1 タイトル" → "8.1.1"
                if let Some(num_str) = trimmed.split_whitespace().next() {
                    let parts: Vec<&str> = num_str.split('.').collect();
                    if parts.len() == 3 && parts.iter().all(|p| p.parse::<u32>().is_ok()) {
                        return Some(num_str.to_string());
                    }
                }
                None
            }
            4 => {
                // "(1) タイトル" → "(1)"
                if trimmed.starts_with('(') {
                    if let Some(end) = trimmed.find(')') {
                        let inner = &trimmed[1..end];
                        if inner.parse::<u32>().is_ok() {
                            return Some(format!("({})", inner));
                        }
                    }
                }
                None
            }
            5 => {
                // "① タイトル" → "①"
                let first_char = trimmed.chars().next()?;
                if is_circled_number(first_char) {
                    return Some(first_char.to_string());
                }
                None
            }
            _ => None,
        }
    }

    fn sync_counters(&mut self, level: u8, number: &str) {
        match level {
            1 => {
                if let Ok(n) = number.parse::<u32>() {
                    self.counters[0] = n;
                    for i in 1..5 {
                        self.counters[i] = 0;
                    }
                }
            }
            2 => {
                let parts: Vec<&str> = number.split('.').collect();
                if parts.len() == 2 {
                    if let (Ok(a), Ok(b)) = (parts[0].parse::<u32>(), parts[1].parse::<u32>()) {
                        self.counters[0] = a;
                        self.counters[1] = b;
                        for i in 2..5 {
                            self.counters[i] = 0;
                        }
                    }
                }
            }
            3 => {
                let parts: Vec<&str> = number.split('.').collect();
                if parts.len() == 3 {
                    if let (Ok(a), Ok(b), Ok(c)) = (
                        parts[0].parse::<u32>(),
                        parts[1].parse::<u32>(),
                        parts[2].parse::<u32>(),
                    ) {
                        self.counters[0] = a;
                        self.counters[1] = b;
                        self.counters[2] = c;
                        for i in 3..5 {
                            self.counters[i] = 0;
                        }
                    }
                }
            }
            4 => {
                let inner = number.trim_start_matches('(').trim_end_matches(')');
                if let Ok(n) = inner.parse::<u32>() {
                    self.counters[3] = n;
                    self.counters[4] = 0;
                }
            }
            5 => {
                if let Some(n) = circled_to_num(number.chars().next().unwrap_or('①')) {
                    self.counters[4] = n;
                }
            }
            _ => {}
        }
    }
}

fn num_to_circled(n: u32) -> String {
    const CIRCLED: &[char] = &[
        '①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩', '⑪', '⑫', '⑬', '⑭', '⑮', '⑯', '⑰', '⑱',
        '⑲', '⑳',
    ];
    if (1..=20).contains(&n) {
        CIRCLED[(n - 1) as usize].to_string()
    } else {
        format!("({})", n)
    }
}

fn is_circled_number(c: char) -> bool {
    ('\u{2460}'..='\u{2473}').contains(&c) // ①-⑳
}

fn circled_to_num(c: char) -> Option<u32> {
    if is_circled_number(c) {
        Some((c as u32) - 0x245F)
    } else {
        None
    }
}
