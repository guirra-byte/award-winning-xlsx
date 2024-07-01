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

struct PodiumProps {
    name: String,
    sales: u64,
}

static mut COMPUTED: Vec<PodiumProps> = vec![];
fn reader(mut root_path: String, file: &str, at: time::Instant) {
    root_path += &"/";
    root_path += file;

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

                                let already_computed = unsafe { &COMPUTED }
                                    .iter()
                                    .enumerate()
                                    .find(|(_, seller)| seller.name == cell.name);

                                match already_computed {
                                    Some((index, data)) => unsafe {
                                        COMPUTED[index].sales += data.sales.clone();
                                    },
                                    None => {
                                        unsafe {
                                            COMPUTED.push(PodiumProps {
                                                name: cell.name.clone(),
                                                sales: cell.amount.clone(),
                                            })
                                        };
                                    }
                                }

                                let target = 55000;
                                if cell.amount >= target {
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
    let dir_content = fs::read_dir(folder.clone());
    let mut computed: Vec<PodiumProps> = vec![];
    match dir_content {
        Ok(dir) => {
            for entry in dir {
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

                                    println!(
                                        "Iniciando verificação dos arquivos: {:?}",
                                        file.start_at
                                    );
                                    reader(file.folder, &file.filename, file.start_at);
                                }
                                None => todo!(),
                            }
                        }
                    }
                    Err(_e) => {
                        eprintln!("Error ao ler o diretório {}", &folder);
                    }
                }
            }
        }
        Err(_) => {
            eprintln!("Não foi possível ler o diretório: {}", &folder);
        }
    }
}

fn main() {
    let filespath = String::from("src/data");
    read_xlsx(filespath);
}
