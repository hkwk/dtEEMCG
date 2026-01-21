use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, anyhow};
use calamine::{Data, Reader, open_workbook_auto};
use chrono::NaiveDateTime;
use regex::Regex;

type DataRow = (
    String,
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
);

fn cell_ref(col_1_based: usize, row_1_based: usize) -> String {
    fn col_to_name(mut col: usize) -> String {
        let mut name = String::new();
        while col > 0 {
            let rem = (col - 1) % 26;
            name.push((b'A' + rem as u8) as char);
            col = (col - 1) / 26;
        }
        name.chars().rev().collect()
    }

    format!("{}{}", col_to_name(col_1_based), row_1_based)
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

fn processed_output_path(input: &Path) -> PathBuf {
    let file_name = input
        .file_name()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "output.xlsx".to_string());
    PathBuf::from(format!("processed_{file_name}"))
}

fn parse_time_to_target_format(time_str: &str) -> Result<String> {
    let time_str = time_str.trim();

    let parsed = if time_str.contains('T') {
        NaiveDateTime::parse_from_str(time_str, "%Y-%m-%dT%H:%M:%S")
            .or_else(|_| NaiveDateTime::parse_from_str(time_str, "%Y-%m-%dT%H:%M:%S%.f"))
    } else if time_str.contains(' ') {
        NaiveDateTime::parse_from_str(time_str, "%Y-%m-%d %H:%M:%S")
            .or_else(|_| NaiveDateTime::parse_from_str(time_str, "%Y/%m/%d %H:%M:%S"))
    } else {
        return Err(anyhow!("无法解析时间格式: {}", time_str));
    };

    let dt = parsed.with_context(|| format!("时间格式错误: {}", time_str))?;

    Ok(dt.format("%Y-%m-%d %H:%M:%S").to_string())
}

fn load_a2_text() -> Result<String> {
    let config_path = Path::new("proton_config.txt");

    if config_path.exists() {
        let content = fs::read_to_string(config_path)
            .with_context(|| format!("无法读取配置文件: {}", config_path.display()))?;
        Ok(content.trim().to_string())
    } else {
        Ok("请参考 proton_config.example.txt 创建配置文件 proton_config.txt".to_string())
    }
}

fn process_excel(path: &Path) -> Result<PathBuf> {
    let mut workbook =
        open_workbook_auto(path).with_context(|| format!("无法打开文件: {}", path.display()))?;

    let sheet_names = workbook.sheet_names();
    let sheet_name = sheet_names
        .first()
        .ok_or_else(|| anyhow!("工作簿中没有工作表"))?;

    let range = workbook
        .worksheet_range(sheet_name)
        .with_context(|| format!("无法读取工作表: {sheet_name}"))?;

    let (height, width) = range.get_size();

    if height < 2 {
        return Err(anyhow!("表格行数不足，无法读取数据"));
    }

    let re = Regex::new(r"\((C|RM)\)").expect("valid regex");

    let mut column_map: std::collections::HashMap<String, usize> = std::collections::HashMap::new();

    for col in 0..width {
        let header = datatype_to_string(range.get((0, col)));
        if !header.is_empty() {
            column_map.insert(header.trim().to_string(), col);
        }
    }

    let time_col = *column_map
        .get("时间")
        .ok_or_else(|| anyhow!("找不到'时间'列"))?;
    let no3_col = *column_map
        .get("NO₃⁻(μg/m³)")
        .ok_or_else(|| anyhow!("找不到'NO₃⁻(μg/m³)'列"))?;
    let so4_col = *column_map
        .get("SO₄²⁻(μg/m³)")
        .ok_or_else(|| anyhow!("找不到'SO₄²⁻(μg/m³)'列"))?;
    let nh4_col = *column_map
        .get("NH₄⁺(μg/m³)")
        .ok_or_else(|| anyhow!("找不到'NH₄⁺(μg/m³)'列"))?;
    let cl_col = *column_map
        .get("Cl⁻(μg/m³)")
        .ok_or_else(|| anyhow!("找不到'Cl⁻(μg/m³)'列"))?;
    let k_col = *column_map
        .get("K⁺(μg/m³)")
        .ok_or_else(|| anyhow!("找不到'K⁺(μg/m³)'列"))?;
    let na_col = *column_map
        .get("Na⁺(μg/m³)")
        .ok_or_else(|| anyhow!("找不到'Na⁺(μg/m³)'列"))?;
    let mg_col = *column_map
        .get("Mg²⁺(μg/m³)")
        .ok_or_else(|| anyhow!("找不到'Mg²⁺(μg/m³)'列"))?;
    let ca_col = *column_map
        .get("Ca²⁺(μg/m³)")
        .ok_or_else(|| anyhow!("找不到'Ca²⁺(μg/m³)'列"))?;

    let mut data_rows: Vec<DataRow> = Vec::new();

    for row in 1..height {
        let time_value = datatype_to_string(range.get((row, time_col)));
        if time_value.is_empty() {
            continue;
        }

        let formatted_time =
            parse_time_to_target_format(&time_value).unwrap_or_else(|_| time_value.clone());

        let is_valid_number = |value: &str| -> bool {
            let value = value.trim();
            if value.is_empty() {
                return false;
            }
            value.parse::<f64>().is_ok()
        };

        let get_value = |col: usize| -> Option<String> {
            let value = datatype_to_string(range.get((row, col)));
            if value.is_empty() || re.is_match(&value) || !is_valid_number(&value) {
                None
            } else {
                Some(value)
            }
        };

        data_rows.push((
            formatted_time,
            get_value(no3_col),
            get_value(so4_col),
            get_value(nh4_col),
            get_value(cl_col),
            get_value(k_col),
            get_value(na_col),
            get_value(mg_col),
            get_value(ca_col),
        ));
    }

    let mut book = umya_spreadsheet::new_file();
    let sheet = book.get_active_sheet_mut();

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

    let mut orange_style = umya_spreadsheet::Style::default();
    orange_style
        .get_fill_mut()
        .get_pattern_fill_mut()
        .set_pattern_type(umya_spreadsheet::structs::PatternValues::Solid);
    orange_style
        .get_fill_mut()
        .get_pattern_fill_mut()
        .get_foreground_color_mut()
        .set_argb("ffff9900");
    orange_style
        .get_fill_mut()
        .get_pattern_fill_mut()
        .get_background_color_mut()
        .set_argb("ffff9900");

    sheet
        .get_cell_mut("A1")
        .set_value("橙色和红色部分请勿改动！！！");
    sheet.get_cell_mut("A1").set_style(red_style.clone());

    let a2_text = load_a2_text()?;
    sheet.get_cell_mut("A2").set_value(a2_text);
    sheet.get_cell_mut("A2").set_style(red_style.clone());

    let row3_headers = [
        "离子色谱",
        "SO₂",
        "HNO₃",
        "HNO₂",
        "HCl",
        "NH₃",
        "NO₃⁻",
        "SO₄²⁻",
        "NH₄⁺",
        "Cl⁻",
        "K⁺",
        "Na⁺",
        "Mg²⁺",
        "Ca²⁺",
        "NO₂⁻",
    ];
    for (i, header) in row3_headers.iter().enumerate() {
        let addr = cell_ref(i + 1, 3);
        sheet.get_cell_mut(addr.as_str()).set_value(*header);
        sheet
            .get_cell_mut(addr.as_str())
            .set_style(orange_style.clone());
    }

    let row4_values = [
        "4401000010003",
        "a21026",
        "a21511",
        "a21510",
        "a21024",
        "a21001",
        "a06006",
        "a06005",
        "a06009",
        "a06008",
        "a06013",
        "a06012",
        "a06011",
        "a06010",
        "a06019",
    ];
    for (i, value) in row4_values.iter().enumerate() {
        let addr = cell_ref(i + 1, 4);
        sheet.get_cell_mut(addr.as_str()).set_value(*value);
        sheet
            .get_cell_mut(addr.as_str())
            .set_style(orange_style.clone());
    }

    let row5_values = [
        "时间", "μg/m³", "μg/m³", "μg/m³", "μg/m³", "μg/m³", "μg/m³", "μg/m³", "μg/m³", "μg/m³",
        "μg/m³", "μg/m³", "μg/m³", "μg/m³", "μg/m³",
    ];
    for (i, value) in row5_values.iter().enumerate() {
        let addr = cell_ref(i + 1, 5);
        sheet.get_cell_mut(addr.as_str()).set_value(*value);
        sheet
            .get_cell_mut(addr.as_str())
            .set_style(orange_style.clone());
    }

    for (row_idx, (time, no3, so4, nh4, cl, k, na, mg, ca)) in data_rows.iter().enumerate() {
        let row = row_idx + 6;

        let time_addr = cell_ref(1, row);
        sheet.get_cell_mut(time_addr.as_str()).set_value(time);
        sheet
            .get_cell_mut(time_addr.as_str())
            .set_style(orange_style.clone());

        let values = [no3, so4, nh4, cl, k, na, mg, ca];
        for (col_idx, value) in values.iter().enumerate() {
            let addr = cell_ref(col_idx + 7, row);
            if let Some(v) = value {
                sheet.get_cell_mut(addr.as_str()).set_value(v);
            } else {
                sheet.get_cell_mut(addr.as_str()).set_value("");
            }
        }
    }

    let output_path = processed_output_path(path);
    umya_spreadsheet::writer::xlsx::write(&book, &output_path)
        .with_context(|| format!("无法保存文件: {}", output_path.display()))?;

    Ok(output_path)
}

pub fn run(args: impl IntoIterator<Item = std::ffi::OsString>) -> Result<()> {
    let mut args = args.into_iter();
    let _exe = args.next();

    let Some(input) = args.next() else {
        println!("请提供文件名作为参数，例如：dtproton proton202552_20260105143932.xlsx");
        return Ok(());
    };

    let input_path = PathBuf::from(input);
    let out = process_excel(&input_path)?;
    println!("文件已处理并保存为: {}", out.display());
    Ok(())
}
