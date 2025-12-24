use std::collections::HashMap;
use std::env;
use std::path::{Path, PathBuf};

use anyhow::{anyhow, Context, Result};
use calamine::{open_workbook_auto, Data, Reader};
use regex::Regex;

#[derive(Debug, Clone)]
struct CellUpdate {
    value: String,
    make_red_fill: bool,
}

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

fn find_target_cells(
    file_path: &Path,
    active_sheet_name: &str,
) -> Result<(usize, usize, HashMap<(u32, u32), CellUpdate>)> {
    let mut workbook = open_workbook_auto(file_path)
        .with_context(|| format!("无法打开文件: {}", file_path.display()))?;

    let range = workbook
        .worksheet_range(active_sheet_name)
        .with_context(|| format!("无法读取工作表: {active_sheet_name}"))?;

    let (height, width) = range.get_size();
    if height == 0 || width == 0 {
        return Ok((height, 0, HashMap::new()));
    }

    // 尽量模拟 Go 版本：
    // maxRow = len(GetRows(activeSheet))（通常到最后一个非空行）
    // maxColumn = max(len(row))（每一行的最后一个非空单元格列号）
    let mut max_column = 0usize;
    for row in 0..height {
        let mut last_non_empty = 0usize;
        for col in 0..width {
            let v = datatype_to_string(range.get((row, col)));
            if !v.is_empty() {
                last_non_empty = col + 1;
            }
        }
        if last_non_empty > max_column {
            max_column = last_non_empty;
        }
    }
    if max_column == 0 {
        max_column = width;
    }

    let re = Regex::new(r"\([^)]*\)").context("无法编译正则表达式")?;

    // 获取第3行特定单元格的值（A1 计数）
    let i3_value = datatype_to_string(range.get((2, 8)));
    let k3_value = datatype_to_string(range.get((2, 10)));
    let q3_value = datatype_to_string(range.get((2, 16)));
    let ay3_value = datatype_to_string(range.get((2, 50)));

    let mut updates: HashMap<(u32, u32), CellUpdate> = HashMap::new();

    for row_1based in 1..=height {
        for col_1based in 1..=max_column {
            let original_value = datatype_to_string(range.get((row_1based - 1, col_1based - 1)));
            let mut value = original_value.clone();
            let mut make_red_fill = false;

            // 替换指定字符串，不设置红色背景
            if value.contains("甲烷非甲烷分析仪") {
                value = value.replace("甲烷非甲烷分析仪", "NMHC监测仪");
            }
            if value.contains("VOCs在线监测仪") {
                value = value.replace("VOCs在线监测仪", "VOCs监测仪");
            }
            if value.contains("总烃(ppbv)") {
                value = value.replace("总烃(ppbv)", "总烃(ppbC)");
            }
            if value.contains("间、对-二甲苯") {
                value = value.replace("间、对-二甲苯", "间/对-二甲苯");
            }
            if value.contains("邻二甲苯") {
                value = value.replace("邻二甲苯", "邻-二甲苯");
            }

            // 新增需求：处理特定列的 -999 替换（从第4行开始）
            if row_1based >= 4 && value.contains("-999") {
                if col_1based == 9 && i3_value == "a24514" {
                    value = "-999#a24041".to_string();
                }
                if col_1based == 11 && k3_value == "a24011" {
                    value = "-999#a24537".to_string();
                }
                if col_1based == 17 && q3_value == "a24510" {
                    value = "-999#a24504".to_string();
                }
                if col_1based == 51 && ay3_value == "a25014" {
                    value = "-999#a25501".to_string();
                }
            }

            // 如果是第3行及之后，删除括号及其中的内容，并设置红色背景
            if row_1based >= 3 && re.is_match(&value) {
                value = re.replace_all(&value, "").to_string();
                make_red_fill = true;
            }

            if value != original_value {
                updates.insert(
                    ((row_1based as u32), (col_1based as u32)),
                    CellUpdate {
                        value: value.trim().to_string(),
                        make_red_fill,
                    },
                );
            }
        }
    }

    Ok((height, max_column, updates))
}

fn process_excel(file_path: &Path) -> Result<()> {
    // 先用 umya 读取，以获取“活动工作表名称”，并在写入前完成工作表重命名。
    let mut book = umya_spreadsheet::reader::xlsx::read(file_path)
        .with_context(|| format!("无法打开文件(写入模式): {}", file_path.display()))?;

    let active_sheet_name_original = book
        .get_active_sheet()
        .get_name()
        .to_string();

    // 重命名工作表（与 Go 版本一致）
    if let Some(sheet) = book.get_sheet_by_name_mut("甲烷非甲烷分析仪") {
        sheet.set_name("NMHC监测仪".to_string());
        println!("工作表名称已从 '甲烷非甲烷分析仪' 替换为 'NMHC监测仪'");
    }
    if let Some(sheet) = book.get_sheet_by_name_mut("VOCs在线监测仪") {
        sheet.set_name("VOCs监测仪".to_string());
        println!("工作表名称已从 'VOCs在线监测仪' 替换为 'VOCs监测仪'");
    }

    // 如果活动表正好被重命名，后续写入时要用新名字；
    // 但 calamine 读取输入文件时仍需要用“旧名字”。
    let active_sheet_name_final = match active_sheet_name_original.as_str() {
        "甲烷非甲烷分析仪" => "NMHC监测仪".to_string(),
        "VOCs在线监测仪" => "VOCs监测仪".to_string(),
        other => other.to_string(),
    };

    let (_max_row, _max_column, updates) =
        find_target_cells(file_path, &active_sheet_name_original)?;

    // 把更新写入到（可能已重命名后的）活动工作表
    let sheet = book
        .get_sheet_by_name_mut(&active_sheet_name_final)
        .ok_or_else(|| anyhow!("找不到工作表: {}", active_sheet_name_final))?;

    // 红色填充样式
    let mut red_style = umya_spreadsheet::Style::default();
    red_style
        .get_fill_mut()
        .get_pattern_fill_mut()
        .set_pattern_type(umya_spreadsheet::structs::PatternValues::Solid);
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

    for (&(row, col), upd) in &updates {
        let addr = to_a1(col, row);
        let cell = sheet.get_cell_mut(addr.as_str());
        cell.set_value(upd.value.as_str());
        if upd.make_red_fill {
            cell.set_style(red_style.clone());
        }
    }

    let base_name = file_path
        .file_name()
        .ok_or_else(|| anyhow!("无法获取文件名"))?
        .to_string_lossy();
    let output_path = PathBuf::from(format!("processed_{base_name}"));
    umya_spreadsheet::writer::xlsx::write(&book, &output_path)
        .with_context(|| format!("无法保存文件: {}", output_path.display()))?;

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
