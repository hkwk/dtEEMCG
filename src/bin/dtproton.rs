fn main() {
    if let Err(e) = dttools::proton::run(std::env::args_os()) {
        eprintln!("处理 Excel 文件时出错: {e:#}");
        std::process::exit(1);
    }
}
