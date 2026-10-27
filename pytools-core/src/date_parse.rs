//! Data-date parsing used by the Rust tools business workflows.

use std::fmt;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CalendarDate {
    year: u16,
    month: u8,
    day: u8,
}

impl CalendarDate {
    pub fn new(year: u16, month: u8, day: u8) -> Option<Self> {
        if year == 0 || !(1..=12).contains(&month) {
            return None;
        }
        let max_day = days_in_month(year, month)?;
        if day == 0 || day > max_day {
            return None;
        }
        Some(Self { year, month, day })
    }

    pub fn year(self) -> u16 {
        self.year
    }

    pub fn month(self) -> u8 {
        self.month
    }

    pub fn day(self) -> u8 {
        self.day
    }
}

impl fmt::Display for CalendarDate {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:04}-{:02}-{:02}", self.year, self.month, self.day)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateValue<'a> {
    Empty,
    Date(CalendarDate),
    DateTime(CalendarDate),
    Text(&'a str),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateSource {
    FileName,
    A1,
}

impl DateSource {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::FileName => "文件名",
            Self::A1 => "A1",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ParsedDataDate {
    pub date: CalendarDate,
    pub source: DateSource,
}

/// Normalize a date value. A year/month string uses the last day of that month.
pub fn normalize_data_date_value(value: DateValue<'_>) -> Option<CalendarDate> {
    match value {
        DateValue::Empty => None,
        DateValue::Date(date) | DateValue::DateTime(date) => Some(date),
        DateValue::Text(text) => parse_text_date(text),
    }
}

/// Prefer a date in the file name, then look for one in the A1 cell text.
pub fn parse_data_date(source_file: &str, sheet_a1: Option<&str>) -> Option<ParsedDataDate> {
    if let Some(date) = parse_text_date(source_file) {
        return Some(ParsedDataDate {
            date,
            source: DateSource::FileName,
        });
    }
    sheet_a1
        .and_then(parse_text_date)
        .map(|date| ParsedDataDate {
            date,
            source: DateSource::A1,
        })
}

fn parse_text_date(text: &str) -> Option<CalendarDate> {
    // Match in the same precedence as the Python implementation: separated
    // full date, compact full date, separated year/month, compact year/month.
    for pattern in [
        Pattern::SeparatedYmd,
        Pattern::CompactYmd,
        Pattern::SeparatedYm,
        Pattern::CompactYm,
    ] {
        if let Some((year, month, day)) = find_first_candidate(text, pattern) {
            let parsed = match day {
                Some(day) => CalendarDate::new(year, month, day),
                None => {
                    days_in_month(year, month).and_then(|day| CalendarDate::new(year, month, day))
                }
            };
            if parsed.is_some() {
                return parsed;
            }
            if matches!(pattern, Pattern::SeparatedYmd | Pattern::CompactYmd) {
                return None;
            }
        }
    }
    None
}

#[derive(Clone, Copy)]
enum Pattern {
    SeparatedYmd,
    CompactYmd,
    SeparatedYm,
    CompactYm,
}

fn find_first_candidate(text: &str, pattern: Pattern) -> Option<(u16, u8, Option<u8>)> {
    let chars: Vec<char> = text.chars().collect();
    for start in 0..chars.len() {
        if start > 0 && is_decimal_digit(chars[start - 1]) {
            continue;
        }
        if chars.get(start) != Some(&'2') || chars.get(start + 1) != Some(&'0') {
            continue;
        }
        let Some(year) = parse_fixed_digits(&chars, start, 4) else {
            continue;
        };
        if year > 2099 {
            continue;
        }
        let tail = start + 4;
        let candidate = match pattern {
            Pattern::SeparatedYmd => parse_separated_ymd(&chars, tail, year),
            Pattern::CompactYmd => parse_compact_ymd(&chars, tail, year),
            Pattern::SeparatedYm => parse_separated_ym(&chars, tail, year),
            Pattern::CompactYm => parse_compact_ym(&chars, tail, year),
        };
        if candidate.is_some() {
            return candidate;
        }
    }
    None
}

fn parse_separated_ymd(chars: &[char], tail: usize, year: u16) -> Option<(u16, u8, Option<u8>)> {
    if !is_date_separator(*chars.get(tail)?) {
        return None;
    }
    let (month, after_month) = parse_one_or_two_digits(chars, tail + 1)?;
    if !is_month_day_separator(*chars.get(after_month)?) {
        return None;
    }
    let (day, mut end) = parse_one_or_two_digits(chars, after_month + 1)?;
    if chars.get(end) == Some(&'日') {
        end += 1;
    }
    if chars.get(end).is_some_and(|ch| is_decimal_digit(*ch)) {
        return None;
    }
    Some((year, month, Some(day)))
}

fn parse_compact_ymd(chars: &[char], tail: usize, year: u16) -> Option<(u16, u8, Option<u8>)> {
    let month = parse_fixed_digits(chars, tail, 2)? as u8;
    let day = parse_fixed_digits(chars, tail + 2, 2)? as u8;
    let end = tail + 4;
    if !(1..=12).contains(&month)
        || !(0..=31).contains(&day)
        || chars.get(end).is_some_and(|ch| is_decimal_digit(*ch))
    {
        return None;
    }
    Some((year, month, Some(day)))
}

fn parse_separated_ym(chars: &[char], tail: usize, year: u16) -> Option<(u16, u8, Option<u8>)> {
    if !is_date_separator(*chars.get(tail)?) {
        return None;
    }
    let (month, mut end) = parse_one_or_two_digits(chars, tail + 1)?;
    if chars.get(end) == Some(&'月') {
        end += 1;
    }
    if chars.get(end).is_some_and(|ch| is_decimal_digit(*ch)) {
        return None;
    }
    Some((year, month, None))
}

fn parse_compact_ym(chars: &[char], tail: usize, year: u16) -> Option<(u16, u8, Option<u8>)> {
    let month = parse_fixed_digits(chars, tail, 2)? as u8;
    if !(1..=12).contains(&month) || chars.get(tail + 2).is_some_and(|ch| is_decimal_digit(*ch)) {
        return None;
    }
    Some((year, month, None))
}

fn parse_fixed_digits(chars: &[char], start: usize, count: usize) -> Option<u16> {
    let mut value = 0_u16;
    for offset in 0..count {
        let digit = chars.get(start + offset)?.to_digit(10)?;
        value = value.checked_mul(10)?.checked_add(digit as u16)?;
    }
    Some(value)
}

fn parse_one_or_two_digits(chars: &[char], start: usize) -> Option<(u8, usize)> {
    let first = chars.get(start)?.to_digit(10)? as u8;
    if let Some(second) = chars.get(start + 1).and_then(|ch| ch.to_digit(10)) {
        Some((first * 10 + second as u8, start + 2))
    } else {
        Some((first, start + 1))
    }
}

fn is_date_separator(ch: char) -> bool {
    matches!(ch, '年' | '-' | '_' | '.' | '/')
}

fn is_month_day_separator(ch: char) -> bool {
    ch == '月' || is_date_separator(ch)
}

fn is_decimal_digit(ch: char) -> bool {
    ch.to_digit(10).is_some()
}

fn days_in_month(year: u16, month: u8) -> Option<u8> {
    Some(match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if is_leap_year(year) => 29,
        2 => 28,
        _ => return None,
    })
}

fn is_leap_year(year: u16) -> bool {
    year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::Value;

    fn calendar_date(value: &str) -> CalendarDate {
        let date = value.get(..10).unwrap_or(value);
        let mut parts = date.split('-');
        let year = parts.next().unwrap().parse().unwrap();
        let month = parts.next().unwrap().parse().unwrap();
        let day = parts.next().unwrap().parse().unwrap();
        assert!(parts.next().is_none());
        CalendarDate::new(year, month, day).unwrap()
    }

    fn normalize_value(value: &Value) -> Option<CalendarDate> {
        match value {
            Value::Null => normalize_data_date_value(DateValue::Empty),
            Value::String(text) => normalize_data_date_value(DateValue::Text(text)),
            Value::Object(object) => {
                let kind = object.get("kind")?.as_str()?;
                let date = calendar_date(object.get("value")?.as_str()?);
                match kind {
                    "date" => normalize_data_date_value(DateValue::Date(date)),
                    "datetime" => normalize_data_date_value(DateValue::DateTime(date)),
                    _ => None,
                }
            }
            _ => None,
        }
    }

    #[test]
    fn matches_shared_python_date_parse_vectors() {
        let vectors: Value =
            serde_json::from_str(include_str!("../tests/fixtures/date_parse_cases.json")).unwrap();

        for case in vectors["normalize"].as_array().unwrap() {
            let id = case["id"].as_str().unwrap();
            let actual = normalize_value(&case["value"]).map(|date| date.to_string());
            let expected = case["expected"].as_str().map(str::to_owned);
            assert_eq!(actual, expected, "normalize case {id}");
        }

        for case in vectors["parse"].as_array().unwrap() {
            let id = case["id"].as_str().unwrap();
            let file_name = case["file_name"].as_str().unwrap();
            let sheet_a1 = case["sheet_a1"].as_str();
            let actual = parse_data_date(file_name, sheet_a1)
                .map(|parsed| (parsed.date.to_string(), parsed.source.as_str().to_owned()));
            let expected = case["expected_date"]
                .as_str()
                .zip(case["expected_source"].as_str())
                .map(|(date, source)| (date.to_owned(), source.to_owned()));
            assert_eq!(actual, expected, "parse case {id}");
        }
    }
}
