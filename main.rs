use calamine::{open_workbook, Reader, Xlsx};
use std::fs;

struct AwardProps {
    name: String,
    amount: u64,
    month: String,
}

fn reader(root_dir: &str, file: &str) {
    let mut filepath = String::from(root_dir);
    filepath += &"/";
    filepath += file;

    // let mut goals_book: Vec<AwardProps> = vec![];
    let mut workbook: Xlsx<_> = open_workbook(filepath).expect("Não foi possível abrir o arquivo.");
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
                                    println!(
                                        "Rust find the winner in month {}, and his name is {}.",
                                        cell.month, cell.name
                                    );
                                }
                            }
                            Err(error) => {
                                eprintln!("{:?} - {:?} - {:?}", error, key, value);
                            }
                        }
                    }
                    // goals_book.push(cell);
                }
            }
            Err(error) => {
                eprintln!("{:?}", error);
            }
        }
    }
}

fn read_xlsx(folder: &str) {
    let dir_content = fs::read_dir(folder).expect("Não foi possível ler a pasta de arquivos.");
    for entry in dir_content {
        match entry {
            Ok(entry) => {
                let (path, name) = (entry.path(), entry.file_name());
                if path.is_file() {
                    match name.to_str() {
                        Some(filename) => {
                            reader(folder, filename);
                        }
                        None => todo!(),
                    }
                }
            }
            Err(e) => {
                eprintln!("Error ao ler o diretório {}", folder);
            }
        }
    }
}

fn main() {
    let filespath = "./data";
    read_xlsx(filespath);
}
