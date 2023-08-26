mod structual;

use pyo3::prelude::*;
use quick_xml::Reader;
use quick_xml::events::Event;
use chrono::NaiveDateTime;
use crate::structual::StructCsv;

#[pyfunction]
fn read_excel(content: Vec<u8>, str_content: Vec<u8>) -> String {
    let name_resolve = str_resolve(str_content);

    // 読み取った内容をXMLとして解析して表示
    let mut xml_reader = Reader::from_reader(&content[..]);
    let mut buffer = Vec::new();
    let mut c_list: Vec<String> = Vec::new();
    let mut row: Vec<StructCsv> = Vec::new();
    let mut is_v = false;
    let mut s: Option<StructCsv> = None;
    let target_attr = vec![115u8, 116u8];  // b"s", b"t"
    let navi = create_navi();
    loop {
        match xml_reader.read_event_into(&mut buffer) {
            Ok(Event::Start(e)) => {
                match e.name().as_ref() {
                    b"c" => {
                        for i in e.attributes() {
                            match i {
                                Ok(x) => {
                                    let a = if target_attr.
                                        contains(&x.key.into_inner()[0]) {
                                        x.key.into_inner()[0].clone()
                                    } else { 0u8 };
                                    s = Some(StructCsv::new(a));
                                }
                                Err(_) => {}
                            }
                        }
                    },
                    b"v" => is_v = true,
                    _ => {},
                }
            }
            Ok(Event::End(e)) => {
                match e.name().as_ref() {
                    b"row" => {
                        let i = row.into_iter().map(|a| {
                            a.clone().get_value(&navi, &name_resolve)
                        }).collect::<Vec<String>>().join(",");
                        c_list.push(i);
                        row = Vec::new();
                    },
                    _ => {},
                }
            }
            Ok(Event::Text(e)) => {
                if is_v {
                    match s {
                        Some(ref mut v) => {
                            let val = e.unescape().unwrap().into_owned();
                            v.set_value(val);
                            row.push(v.clone());
                            s = None;
                        }
                        None => {}
                    }
                    is_v = false;
                }
            }
            Ok(Event::Eof) => {
                break;
            }
            Err(e) => {
                eprintln!("Error: {}", e);
                break;
            }
            _ => {}
        }
    }
    c_list.join("\n")
}

fn str_resolve(content: Vec<u8>) ->  Vec<String> {
    // 読み取った内容をXMLとして解析して表示
    let mut xml_reader = Reader::from_reader(&content[..]);
    let mut buffer = Vec::new();
    let mut name_resolve: Vec<String> = Vec::new();
    let mut is_text = false;
    let mut no_text_ = false;
    loop {
        match xml_reader.read_event_into(&mut buffer) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"t" => is_text = true,
                    b"rPh" => no_text_ = true,
                    _ => no_text_ = false,
                }
            }
            Ok(Event::Text(e)) => {
                if &is_text & !&no_text_ {
                    let val = e.unescape().unwrap().into_owned();
                    name_resolve.push(val);
                    is_text = false;
                    no_text_ = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                eprintln!("Error: {}", e);
                break;
            }
            _ => {}
        }
    }
    name_resolve
}

fn create_navi() -> NaiveDateTime {
    NaiveDateTime::new(
        chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap(),
        chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap())
}

/// A Python module implemented in Rust.
#[pymodule]
fn excel2csv(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(read_excel, m)?)?;
    Ok(())
}
