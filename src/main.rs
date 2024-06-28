use calamine::{open_workbook, Reader, Xlsx};
use std::fs::{self};
use std::time;

struct FileProps {
    folder: String,
    filename: String,
    start_at: time::Instant,
}
struct AwardProps {
    name: String,
    amount: u64,
    month: String,
}

fn reader(mut root_path: String, file: &str, at: time::Instant) {
    root_path += &"/";
    root_path += file;

    // let mut goals_book: Vec<AwardProps> = vec![];
    let mut workbook: Xlsx<_> =
        open_workbook(root_path).expect("Não foi possível abrir o arquivo.");
    for sheet_name in workbook.sheet_names().to_owned() {
        let worksheet = workbook.worksheet_range(&sheet_name);
        match worksheet {
            Ok(data) => {
                for row in data.rows() {
                    let (key, value) = (&row[0], &row[1]);
                    if key != "Vendedor" && value != "Venda" {
                        let amount = value.to_string().parse::<u64>();
                        match amount {
                            Ok(result) => {
                                let cell: AwardProps = AwardProps {
                                    name: String::from(key.to_string()),
                                    amount: result,
                                    month: file.to_string(),
                                };

                                let winner = 55000;
                                if cell.amount >= winner {
                                    let execution = at.elapsed();
                                    println!(
                                        "Rust find the winner in month {}, and his name is {} on \n At {:?}",
                                        cell.month, cell.name, execution
                                    );
                                }
                            }
                            Err(error) => {
                                eprintln!("{:?} - {:?} - {:?}", error, key, value);
                            }
                        }
                    }
                }
            }
            Err(error) => {
                eprintln!("{:?}", error);
            }
        }
    }
}

fn read_xlsx(folder: String) {
    let dir_content =
        fs::read_dir(folder.clone()).expect("Não foi possível ler a pasta de arquivos.");
    for entry in dir_content {
        match entry {
            Ok(entry) => {
                let (path, name) = (entry.path(), entry.file_name());
                if path.is_file() {
                    match name.to_str() {
                        Some(filename) => {
                            let start = time::Instant::now();
                            let file: FileProps = FileProps {
                                filename: filename.to_string(),
                                folder: folder.clone(),
                                start_at: start,
                            };

                            println!("Iniciando verificação dos arquivos: {:?}", file.start_at);
                            reader(file.folder, &file.filename, file.start_at);
                        }
                        None => todo!(),
                    }
                }
            }
            Err(_e) => {
                eprintln!("Error ao ler o diretório {}", folder.clone());
            }
        }
    }
}

fn main() {
    let filespath = "src/data";
    read_xlsx(filespath.to_string());
}
