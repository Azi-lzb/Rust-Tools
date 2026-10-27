//! Rule-based path and column-spec handling used by tools workflows.

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ColumnSpecValue<'a> {
    Missing,
    Text(&'a str),
    Integer(i64),
    Number(f64),
}

pub fn match_all_keywords(source: &str, keywords: &str) -> bool {
    if keywords.is_empty() {
        return true;
    }
    let source_lower = source.to_lowercase();
    split_keywords(keywords)
        .into_iter()
        .all(|keyword| source_lower.contains(keyword.to_lowercase().as_str()))
}

/// Parse the scalar forms accepted by the original Python configuration reader.
pub fn parse_col_spec_value(value: ColumnSpecValue<'_>) -> usize {
    match value {
        ColumnSpecValue::Missing => 0,
        ColumnSpecValue::Text(text) => parse_col_spec(text),
        ColumnSpecValue::Integer(number) => parse_col_spec(&number.to_string()),
        ColumnSpecValue::Number(number) if number.is_nan() => 0,
        ColumnSpecValue::Number(number) => parse_col_spec(&number.to_string()),
    }
}

pub fn parse_col_spec(value: &str) -> usize {
    let text = value.trim();
    if text.is_empty() {
        return 0;
    }

    if text.chars().all(|ch| ch.is_ascii_digit()) {
        return text.parse().unwrap_or(0);
    }
    if text.bytes().filter(|byte| *byte == b'.').count() == 1
        && text
            .bytes()
            .all(|byte| byte.is_ascii_digit() || byte == b'.')
    {
        return text
            .parse::<f64>()
            .map(|number| number as usize)
            .unwrap_or(0);
    }

    let mut column = 0_usize;
    for byte in text.bytes() {
        let upper = byte.to_ascii_uppercase();
        if !upper.is_ascii_uppercase() {
            return 0;
        }
        column = match column
            .checked_mul(26)
            .and_then(|number| number.checked_add((upper - b'A' + 1) as usize))
        {
            Some(value) => value,
            None => return 0,
        };
    }
    column
}

/// Parse a single scalar configuration value as a list of unique column numbers.
pub fn parse_col_specs_value(value: ColumnSpecValue<'_>) -> Vec<usize> {
    match value {
        ColumnSpecValue::Missing => Vec::new(),
        ColumnSpecValue::Text(text) => parse_col_specs(text),
        ColumnSpecValue::Integer(number) => parse_col_specs(&number.to_string()),
        ColumnSpecValue::Number(number) if number.is_nan() => Vec::new(),
        ColumnSpecValue::Number(number) => parse_col_specs(&number.to_string()),
    }
}

pub fn parse_col_specs(value: &str) -> Vec<usize> {
    let normalized = value
        .replace('，', ",")
        .replace('；', ",")
        .replace(';', ",");
    let mut columns = Vec::new();
    for part in normalized.split(',') {
        let column = parse_col_spec(part);
        if column > 0 && !columns.contains(&column) {
            columns.push(column);
        }
    }
    columns
}

fn split_keywords(value: &str) -> Vec<String> {
    value
        .replace('，', ",")
        .replace('；', ";")
        .replace(',', ";")
        .split(';')
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .map(str::to_owned)
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::Value;

    fn column_spec_value(value: &Value) -> ColumnSpecValue<'_> {
        match value {
            Value::Null => ColumnSpecValue::Missing,
            Value::String(text) => ColumnSpecValue::Text(text),
            Value::Object(object) => match object.get("kind").and_then(Value::as_str) {
                Some("integer") => ColumnSpecValue::Integer(object["value"].as_i64().unwrap()),
                Some("number") => ColumnSpecValue::Number(object["value"].as_f64().unwrap()),
                Some("nan") => ColumnSpecValue::Number(f64::NAN),
                kind => panic!("unsupported column input kind: {kind:?}"),
            },
            other => panic!("unsupported column input: {other}"),
        }
    }

    #[test]
    fn matches_shared_python_keyword_and_column_vectors() {
        let vectors: Value =
            serde_json::from_str(include_str!("../tests/fixtures/path_map_cases.json")).unwrap();

        for case in vectors["keywords"].as_array().unwrap() {
            let actual = match_all_keywords(
                case["source"].as_str().unwrap(),
                case["keywords"].as_str().unwrap(),
            );
            assert_eq!(
                actual,
                case["expected"].as_bool().unwrap(),
                "keyword case {}",
                case["id"].as_str().unwrap()
            );
        }

        for case in vectors["columns_single"].as_array().unwrap() {
            let actual = parse_col_spec_value(column_spec_value(&case["value"]));
            assert_eq!(
                actual,
                case["expected"].as_u64().unwrap() as usize,
                "single-column case {}",
                case["id"].as_str().unwrap()
            );
        }

        for case in vectors["columns_multiple"].as_array().unwrap() {
            let actual = parse_col_specs_value(column_spec_value(&case["value"]));
            let expected = case["expected"]
                .as_array()
                .unwrap()
                .iter()
                .map(|value| value.as_u64().unwrap() as usize)
                .collect::<Vec<_>>();
            assert_eq!(
                actual,
                expected,
                "multiple-column case {}",
                case["id"].as_str().unwrap()
            );
        }
    }
}
