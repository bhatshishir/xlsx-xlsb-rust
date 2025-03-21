use notify::{Watcher, RecursiveMode, RecommendedWatcher, Result, Event, EventKind};
use std::path::{Path, PathBuf};
use std::sync::mpsc::channel;
use std::fs;
use std::process::Command;

fn main() -> Result<()> {
    let folder_path = "C:\\ExcelDrop";  // Folder you’ll drop Excel files into

    // Create the folder if it doesn't exist
    if !Path::new(folder_path).exists() {
        fs::create_dir(folder_path)?;
        println!("Created folder: {}", folder_path);
    }

    println!("🔍 Watching folder: {}", folder_path);

    // Create channel and watcher
    let (tx, rx) = channel();
    let mut watcher: RecommendedWatcher = notify::recommended_watcher(tx)?;
    watcher.watch(Path::new(folder_path), RecursiveMode::NonRecursive)?;

    // Listen for new files
    for event in rx {
        match event {
            Ok(Event { kind: EventKind::Create(_), paths, .. }) => {
                for path in paths {
                    if let Some(ext) = path.extension() {
                        if (ext == "xlsx" || ext == "xls") && !path.file_name().unwrap().to_string_lossy().starts_with("~$") {                    
                            println!("📁 New file detected: {:?}", path);
                            convert_to_xlsb(&path);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(())
}

fn convert_to_xlsb(file_path: &PathBuf) {
    let input = file_path.to_str().unwrap();

    // Replace the extension with .xlsb safely
    let mut output_path = file_path.clone();
    output_path.set_extension("xlsb");
    let output = output_path.to_str().unwrap();

    println!("🔄 Converting: {} → {}", input, output);

    let result = Command::new("python")
        .arg("-c")
        .arg(format!(
            "import win32com.client as win32; excel = win32.gencache.EnsureDispatch('Excel.Application'); wb = excel.Workbooks.Open(r'{}'); wb.SaveAs(r'{}', FileFormat=50); wb.Close(); excel.Quit()",
            input, output
        ))
        .output();

    match result {
        Ok(out) => {
            if out.status.success() {
                println!("✅ Successfully converted to XLSB: {}", output);
            } else {
                eprintln!("❌ Conversion failed: {}", String::from_utf8_lossy(&out.stderr));
            }
        }
        Err(e) => eprintln!("❌ Failed to run Python: {}", e),
    }
}
