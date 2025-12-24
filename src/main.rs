use std::collections::HashMap;
use std::env;
use std::path::{Path, PathBuf};

use anyhow::{anyhow, Context, Result};
use calamine::{open_workbook_auto, Data, Reader};
use regex::Regex;

fn datatype_to_string(cell: Option<&Data>) -> String {
    match cell {
        None => String::new(),
        Some(Data::Empty) => String::new(),
        Some(Data::String(s)) => s.clone(),
        Some(Data::Float(n)) => {
            if n.fract() == 0.0 {
                format!("{:.0}", n)
            } else {
                n.to_string()
            }
        }
        Some(Data::Int(n)) => n.to_string(),
        Some(Data::Bool(b)) => b.to_string(),
        Some(Data::Error(e)) => format!("{e:?}"),
        Some(Data::DateTime(f)) => f.to_string(),
        Some(other) => format!("{other:?}"),
    }
}

fn column_number_to_name(mut column: u32) -> String {
    // 1 -> A, 26 -> Z, 27 -> AA ...
    let mut name = String::new();
    while column > 0 {
        let rem = ((column - 1) % 26) as u8;
        name.insert(0, (b'A' + rem) as char);
        column = (column - 1) / 26;
    }
    name
}

fn to_a1(col_1based: u32, row_1based: u32) -> String {
    format!("{}{}", column_number_to_name(col_1based), row_1based)
}

fn find_target_cells(file_path: &Path) -> Result<(String, usize, usize, HashMap<(u32, u32), String>)> {
    let mut workbook = open_workbook_auto(file_path)
        .with_context(|| format!("无法打开文件: {}", file_path.display()))?;

    let sheet_name = workbook
        .sheet_names()
        .get(0)
        .cloned()
        .ok_or_else(|| anyhow!("工作簿中没有工作表"))?;

    let range = workbook
        .worksheet_range(&sheet_name)
        .with_context(|| format!("无法读取工作表: {sheet_name}"))?;

    let (height, width) = range.get_size();
    if height == 0 || width == 0 {
        return Ok((sheet_name, height, 0, HashMap::new()));
    }

    // 尽量模拟 Go 版本：maxRow=len(rows)，maxColumn=len(rows[0])
    // calamine 的 range 宽度可能更大，因此这里优先以第一行“最后一个非空单元格”作为 maxColumn。
    let mut max_column = 0usize;
    for col in 0..width {
        let v = datatype_to_string(range.get((0, col)));
        if !v.is_empty() {
            max_column = col + 1;
        }
    }
    if max_column == 0 {
        max_column = width;
    }

    let re = Regex::new(r"\([^)]*\)").context("无法编译正则表达式")?;
    let mut updates: HashMap<(u32, u32), String> = HashMap::new();

    // 从第三行开始（1-based） => 0-based row index 从 2 开始
    for row in 2..height {
        for col in 0..max_column {
            let value = datatype_to_string(range.get((row, col)));
            if re.is_match(&value) {
                let new_value = re.replace_all(&value, "").to_string();
                // umya 的坐标是 1-based
                updates.insert(((row + 1) as u32, (col + 1) as u32), new_value);
            }
        }
    }

    Ok((sheet_name, height, max_column, updates))
}

fn apply_updates_and_save(
    file_path: &Path,
    sheet_name: &str,
    updates: &HashMap<(u32, u32), String>,
) -> Result<PathBuf> {
    let mut book = umya_spreadsheet::reader::xlsx::read(file_path)
        .with_context(|| format!("无法打开文件(写入模式): {}", file_path.display()))?;

    let sheet = book
        .get_sheet_by_name_mut(sheet_name)
        .ok_or_else(|| anyhow!("找不到工作表: {sheet_name}"))?;

    // 红色填充样式
    let mut red_style = umya_spreadsheet::Style::default();
    red_style
        .get_fill_mut()
        .get_pattern_fill_mut()
        .set_pattern_type(umya_spreadsheet::structs::PatternValues::Solid);
    // 注意：umya-spreadsheet 的 Color::set_argb() 若传入值刚好等于它的内置 INDEXED_COLORS，
    // 会自动转换为 indexed="n"，这在不同查看器/调色板下可能显示成非预期颜色（比如绿色）。
    // 为确保写入的是 rgb="..."，这里使用小写形式避免命中该映射；Excel 对 hex 大小写不敏感。
    // ARGB: FF + RRGGBB => #FF0000
    red_style
        .get_fill_mut()
        .get_pattern_fill_mut()
        .get_foreground_color_mut()
        .set_argb("ffff0000");
    red_style
        .get_fill_mut()
        .get_pattern_fill_mut()
        .get_background_color_mut()
        .set_argb("ffff0000");

    for (&(row, col), new_value) in updates {
        let addr = to_a1(col, row);
        let cell = sheet.get_cell_mut(addr.as_str());
        cell.set_value(new_value);
        cell.set_style(red_style.clone());
    }

    let base_name = file_path
        .file_name()
        .ok_or_else(|| anyhow!("无法获取文件名"))?
        .to_string_lossy();
    let output_path = PathBuf::from(format!("processed_{base_name}"));

    umya_spreadsheet::writer::xlsx::write(&book, &output_path)
        .with_context(|| format!("无法保存文件: {}", output_path.display()))?;

    Ok(output_path)
}

fn process_excel(file_path: &Path) -> Result<()> {
    let (sheet_name, _max_row, _max_column, updates) = find_target_cells(file_path)?;
    let output_path = apply_updates_and_save(file_path, &sheet_name, &updates)?;
    println!("文件已处理并保存为: {}", output_path.display());
    Ok(())
}

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("请提供文件名作为参数，例如：./program 45vocs2.xlsx");
        return;
    }

    let file_path = Path::new(&args[1]);
    if let Err(err) = process_excel(file_path) {
        eprintln!("处理Excel文件时出错: {err:#}");
    }
}
