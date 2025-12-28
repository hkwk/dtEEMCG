fn main() {
    if let Err(err) = dttools::eemcg::run(std::env::args_os()) {
        eprintln!("处理Excel文件时出错: {err:#}");
        std::process::exit(1);
    }
}
