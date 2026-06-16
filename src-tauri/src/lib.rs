use serde::Serialize;
use std::fs;
use std::path::PathBuf;
use std::process::Command;

// 嵌入 grid_preview.py 的内容
const GRID_PREVIEW_PY: &str = include_str!("../grid_preview.py");

#[derive(Debug, Serialize, Clone)]
struct DocxFile {
    name: String,
    size: u64,
    modified: String,
    has_html: bool,
}

fn exe_dir() -> PathBuf {
    std::env::current_exe()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf()
}

fn docx_dir() -> PathBuf {
    exe_dir().join("docx")
}

fn html_dir() -> PathBuf {
    exe_dir().join("html")
}

fn py_script_path() -> PathBuf {
    exe_dir().join("grid_preview.py")
}

fn ensure_dirs_and_script() {
    let d = docx_dir();
    let h = html_dir();
    let py = py_script_path();

    let _ = fs::create_dir_all(&d);
    let _ = fs::create_dir_all(&h);

    // 如果 grid_preview.py 不存在或内容过旧，写入最新版
    if !py.exists() {
        let _ = fs::write(&py, GRID_PREVIEW_PY);
    }
}

fn format_time(secs: u64) -> String {
    let days = secs / 86400;
    if days > 0 { return format!("{} 天前", days); }
    let hours = secs / 3600;
    if hours > 0 { return format!("{} 小时前", hours); }
    let mins = secs / 60;
    format!("{} 分钟前", mins.max(1))
}

#[tauri::command]
fn scan_docx_dir() -> Vec<DocxFile> {
    ensure_dirs_and_script();
    let d = docx_dir();
    let h = html_dir();

    let mut files = Vec::new();
    if let Ok(entries) = fs::read_dir(&d) {
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().map(|e| e == "docx").unwrap_or(false) {
                let name = path.file_name().unwrap().to_string_lossy().to_string();
                let metadata = path.metadata().unwrap_or_else(|_| {
                    path.symlink_metadata().unwrap()
                });
                let size = metadata.len();
                let modified = metadata
                    .modified()
                    .map(|t| {
                        t.elapsed()
                            .map(|d| format_time(d.as_secs()))
                            .unwrap_or_default()
                    })
                    .unwrap_or_default();

                let html_name = name.replace(".docx", "_方格纸预览.html");
                let has_html = h.join(&html_name).exists();

                files.push(DocxFile { name, size, modified, has_html });
            }
        }
    }
    files.sort_by(|a, b| b.name.cmp(&a.name));
    files
}

#[tauri::command]
fn import_docx(source: String) -> Result<Vec<DocxFile>, String> {
    ensure_dirs_and_script();
    let src = PathBuf::from(&source);
    if !src.exists() {
        return Err(format!("文件不存在: {}", source));
    }
    let fname = src.file_name().unwrap().to_string_lossy().to_string();
    if !fname.ends_with(".docx") {
        return Err("仅支持 .docx 文件".into());
    }
    let dest = docx_dir().join(&fname);
    fs::copy(&src, &dest).map_err(|e| format!("复制失败: {}", e))?;
    Ok(scan_docx_dir())
}

#[tauri::command]
fn generate_preview(filename: String) -> Result<String, String> {
    ensure_dirs_and_script();
    let docx_path = docx_dir().join(&filename);
    if !docx_path.exists() {
        return Err(format!("文件不存在: {}", filename));
    }
    let py = py_script_path();
    if !py.exists() {
        let _ = fs::write(&py, GRID_PREVIEW_PY);
    }

    let d_str = docx_dir().to_string_lossy().to_string();
    let h_str = html_dir().to_string_lossy().to_string();

    let output = Command::new("python")
        .arg(py.to_string_lossy().to_string())
        .arg(&d_str)
        .arg(&h_str)
        .output()
        .map_err(|e| format!("运行 Python 失败: {}。请确保已安装 Python 并加入 PATH", e))?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        let stdout = String::from_utf8_lossy(&output.stdout);
        return Err(format!("生成失败:\n{}", if stderr.is_empty() { stdout } else { stderr }));
    }

    let html_name = filename.replace(".docx", "_方格纸预览.html");
    let html_path = html_dir().join(&html_name);
    if html_path.exists() {
        Ok(html_path.to_string_lossy().to_string())
    } else {
        let stdout = String::from_utf8_lossy(&output.stdout);
        Err(format!("HTML 未生成。输出:\n{}", stdout))
    }
}

#[tauri::command]
fn get_html_path(filename: String) -> Result<String, String> {
    let html_name = filename.replace(".docx", "_方格纸预览.html");
    let path = html_dir().join(&html_name);
    if path.exists() {
        Ok(path.to_string_lossy().to_string())
    } else {
        Err("预览文件不存在，请先生成".into())
    }
}

#[tauri::command]
fn read_html_content(filename: String) -> Result<String, String> {
    let html_name = filename.replace(".docx", "_方格纸预览.html");
    let path = html_dir().join(&html_name);
    if !path.exists() {
        return Err("预览文件不存在，请先生成".into());
    }
    fs::read_to_string(&path).map_err(|e| format!("读取失败: {}", e))
}

#[tauri::command]
fn open_html_in_browser(filename: String) -> Result<(), String> {
    let html_name = filename.replace(".docx", "_方格纸预览.html");
    let path = html_dir().join(&html_name);
    if !path.exists() {
        return Err("预览文件不存在".into());
    }
    let url = format!("file:///{}", path.to_string_lossy().replace('\\', "/"));
    open::that(&url).map_err(|e| format!("打开失败: {}", e))
}

#[tauri::command]
async fn pick_docx_file() -> Result<Option<Vec<String>>, String> {
    let files = rfd::AsyncFileDialog::new()
        .add_filter("Word 文档", &["docx"])
        .set_title("选择 .docx 文件")
        .pick_files()
        .await;

    match files {
        Some(paths) => {
            let result: Vec<String> = paths.iter()
                .map(|p| p.path().to_string_lossy().to_string())
                .collect();
            Ok(Some(result))
        }
        None => Ok(None),
    }
}

#[tauri::command]
fn win_minimize(window: tauri::Window) {
    let _ = window.minimize();
}

#[tauri::command]
fn win_toggle_maximize(window: tauri::Window) {
    if window.is_maximized().unwrap_or(false) {
        let _ = window.unmaximize();
    } else {
        let _ = window.maximize();
    }
}

#[tauri::command]
fn win_close(window: tauri::Window) {
    let _ = window.close();
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_shell::init())
        .setup(|_app| {
            ensure_dirs_and_script();
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            scan_docx_dir,
            import_docx,
            generate_preview,
            get_html_path,
            read_html_content,
            open_html_in_browser,
            pick_docx_file,
            win_minimize,
            win_toggle_maximize,
            win_close,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
