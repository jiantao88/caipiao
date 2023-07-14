use reqwest::Error as ReqwestError;
use regex::Regex;
use futures::stream::{self, StreamExt};
use std::error::Error;
use std::sync::{Arc, Mutex};
use xlsxwriter::*;

#[tokio::main]
async fn main() -> Result<(), Box<dyn Error>> {
    let client = reqwest::Client::new();
    let total_pages = 5;
    let base_url = "http://kaijiang.zhcw.com/zhcw/html/ssq/list_{}.html";

    let re = Regex::new(r#"(?s)<tr>.*?<td align=\"center\">(.*?)</td>.*?<td align=\"center\">(.*?)</td>.*?<td align=\"center\" style=\"padding-left:10px;\">.*?<em class=\"rr\">(.*?)</em>.*?<em class=\"rr\">(.*?)</em>.*?<em class=\"rr\">(.*?)</em>.*?<em class=\"rr\">(.*?)</em>.*?<em class=\"rr\">(.*?)</em>.*?<em class=\"rr\">(.*?)</em>.*?<em>(.*?)</em></td>"#).unwrap();

    let workbook = Workbook::new("双色球统计结果.xlsx")?;
    let mut worksheet = workbook.add_worksheet(None)?;

    worksheet.write_string(0, 0, "日期", None)?;
    worksheet.write_string(0, 1, "期数", None)?;
    worksheet.write_string(0, 2, "第一个红球", None)?;
    worksheet.write_string(0, 3, "第二个红球", None)?;
    worksheet.write_string(0, 4, "第三个红球", None)?;
    worksheet.write_string(0, 5, "第四个红球", None)?;
    worksheet.write_string(0, 6, "第五个红球", None)?;
    worksheet.write_string(0, 7, "第六个红球", None)?;
    worksheet.write_string(0, 8, "蓝球", None)?;

    let pages: Vec<_> = (1..=total_pages).collect();
    let results = Arc::new(Mutex::new(Vec::new()));

    let bodies = stream::iter(pages)
        .map(|page_num| {
            let client = &client;
            let results = Arc::clone(&results);
            let url = format!("{}", base_url.replace("{}", &page_num.to_string()));
            println!("url={}", url);
            async move {
                let resp = client.get(&url).send().await.map_err(ReqwestError::from)?;
                let body = resp.text().await.map_err(ReqwestError::from)?;
                // println!("body={}", body);
                Ok::<_, ReqwestError>((body, page_num, results))
            }
        })
        .buffer_unordered(total_pages);

    bodies.for_each(|result| async {
        match result {
            Ok((body, page_num, results)) => {
                println!("正在写入第{}页", page_num);
                // println!("body={}", &body);
                // println!("cap={:?}", re.captures_iter(&body));
               let mut _capt= re.captures_iter(&body);
               
                if _capt.next().is_none() {
                    println!("No matches found");
                }else {
                    for cap in _capt {
                        let row: Vec<String> = cap.iter().skip(1).map(|m| m.unwrap().as_str().to_string()).collect();
                        println!("row={:?}", row);
                        results.lock().unwrap().push(row);
                    }
                }
                
            }
            Err(e) => eprintln!("Error: {}", e),
        }
    }).await;

    let results = Arc::try_unwrap(results).unwrap().into_inner().unwrap();

    for (i, row) in results.iter().enumerate() {
        for (j, item) in row.iter().enumerate() {
            worksheet.write_string((i+1) as u32, j as u16, item, None).unwrap();
        }
    }

    workbook.close().unwrap();

    Ok(())
}
