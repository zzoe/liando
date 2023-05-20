use std::collections::HashMap;
use std::path::Path;

use rfd::AsyncFileDialog;
use serde::{Deserialize, Serialize};
use slint::SharedString;
use time::{macros::format_description, Date, Duration};
use umya_spreadsheet::helper::coordinate::{coordinate_from_index, string_from_column_index};
use umya_spreadsheet::{HorizontalAlignmentValues, Style, VerticalAlignmentValues};

use crate::App;

pub(crate) async fn execute_handle(ui_weak: slint::Weak<App>) {
    let (s, r) = async_channel::bounded(1);
    let ui_weak_copy = ui_weak.clone();
    ui_weak_copy
        .upgrade_in_event_loop(move |ui| {
            let (start_date, end_date) = (ui.get_start_date(), ui.get_end_date());
            let mut req = ReportReq {
                start_date: start_date.to_string(),
                end_date: end_date.to_string(),
                statistics_file: ui.get_statistics_file().to_string(),
                record_file: ui.get_record_file().to_string(),
            };

            let (start_date_check, end_date_check) = req.check_date();
            if start_date_check && end_date_check {
                if start_date > end_date {
                    ui.set_start_date(end_date);
                    ui.set_end_date(start_date);
                    std::mem::swap(&mut req.start_date, &mut req.end_date);
                }
                s.send_blocking(Some(req)).unwrap();
            } else {
                let mut text = SharedString::from("输入有误，请检查!");
                if !start_date_check {
                    text = SharedString::from("开始日期有误!");
                } else if !end_date_check {
                    text = SharedString::from("结束日期有误!");
                }

                ui.set_error_text(text);
                ui.invoke_alert_error();
            }

            s.close();
        })
        .unwrap();

    if let Some(report_req) = r.recv().await.unwrap_or_default() {
        if let Err(e) = generate_report(report_req).await {
            ui_weak_copy
                .upgrade_in_event_loop(move |ui| {
                    ui.set_error_text(e.to_string().into());
                    ui.invoke_alert_error();
                })
                .unwrap();
        }
    }

    ui_weak_copy
        .upgrade_in_event_loop(move |ui| {
            ui.set_button_enabled(true);
        })
        .unwrap();
}

#[derive(Default, Deserialize, Serialize)]
struct ReportReq {
    start_date: String,
    end_date: String,
    statistics_file: String,
    record_file: String,
}

impl ReportReq {
    fn check_date(&self) -> (bool, bool) {
        let format = format_description!("[year]-[month]-[day]");
        let start_date = Date::parse(&self.start_date, &format);
        let end_date = Date::parse(&self.end_date, &format);

        (start_date.is_ok(), end_date.is_ok())
    }

    fn get_start_date(&self) -> Date {
        let format = format_description!("[year]-[month]-[day]");
        Date::parse(&self.start_date, &format).unwrap()
    }

    fn get_end_date(&self) -> Date {
        let format = format_description!("[year]-[month]-[day]");
        Date::parse(&self.end_date, &format).unwrap()
    }
}

async fn generate_report(req: ReportReq) -> anyhow::Result<()> {
    let mut dates = HashMap::new();
    let mut loop_date = req.get_start_date();
    let end_date = req.get_end_date();
    let format = format_description!("[year repr:last_two]-[month]-[day]");
    while loop_date <= end_date {
        dates.insert(loop_date.format(&format).unwrap(), loop_date);
        loop_date = loop_date.saturating_add(Duration::days(1));
    }

    //从每日统计文件取每人每天上下班的考勤状态(G,I)和工作时长（Q）
    let mut all_attendance = get_everyone_statistics(Path::new(&req.statistics_file), &dates)?;
    //从原始记录文件取每人每天上下班的虚拟打卡情况(N)
    get_everyone_abnormal_reason(Path::new(&req.record_file), &dates, &mut all_attendance)?;
    //输出文件
    save_res(dates, all_attendance).await?;

    /*
       1. 默认Sheet取的第一页
       2. 当天上/下班既有正常打卡，也有虚拟打卡，是否记虚拟?
       3. 各种状态的映射关系?

       可配置项：每个值的列、打卡结果映射成枚举
    */
    Ok(())
}

struct Attendance {
    employee_name: String,
    enter_info: String,
    leave_info: String,
    work_hours: u32,
    abnormal_reason: String,
}

// HashMap<工号，HashMap<日期， 考勤情况>>
fn get_everyone_statistics(
    path: impl AsRef<Path>,
    dates: &HashMap<String, Date>,
) -> anyhow::Result<HashMap<String, HashMap<Date, Attendance>>> {
    let book = umya_spreadsheet::reader::xlsx::read(path.as_ref())?;
    let worksheet = book.get_sheet(&0).map_err(|e| anyhow::anyhow!(e))?;
    let (_, max_row) = worksheet.get_highest_column_and_row();
    let mut res = HashMap::new();

    for r in 5..max_row + 1 {
        if let Some(attendance_date) = worksheet.get_formatted_value((5, r)).get(..8) {
            // println!("{attendance_date}");
            if let Some(date) = dates.get(attendance_date) {
                let employee_id = worksheet.get_formatted_value((2, r));
                if employee_id.is_empty() {
                    continue;
                }
                if !res.contains_key(&employee_id) {
                    res.insert(employee_id.to_string(), HashMap::new());
                }

                let employee_attendance = res.get_mut(&employee_id).unwrap();
                employee_attendance.insert(
                    *date,
                    Attendance {
                        employee_name: worksheet.get_formatted_value((1, r)),
                        enter_info: worksheet.get_formatted_value((7, r)),
                        leave_info: worksheet.get_formatted_value((9, r)),
                        abnormal_reason: String::new(),
                        work_hours: worksheet.get_value_number((17, r)).unwrap_or_default() as u32,
                    },
                );
            }
        }
    }

    Ok(res)
}

fn get_everyone_abnormal_reason(
    path: impl AsRef<Path>,
    dates: &HashMap<String, Date>,
    res: &mut HashMap<String, HashMap<Date, Attendance>>,
) -> anyhow::Result<()> {
    let book = umya_spreadsheet::reader::xlsx::read(path.as_ref())?;
    let worksheet = book.get_sheet(&0).map_err(|e| anyhow::anyhow!(e))?;
    let (_, max_row) = worksheet.get_highest_column_and_row();

    for r in 4..max_row + 1 {
        let employee_id = worksheet.get_formatted_value((4, r));
        // 考勤里面有这个工号
        if let Some(employee_attendance) = res.get_mut(&employee_id) {
            if let Some(date) = worksheet.get_formatted_value((7, r)).get(..8) {
                // 查询范围包含这一天
                if let Some(attendance_date) = dates.get(date) {
                    // 考勤里面有这一天
                    if let Some(attendance) = employee_attendance.get_mut(attendance_date) {
                        attendance.abnormal_reason = worksheet.get_formatted_value((14, r));
                    }
                }
            }
        }
    }

    Ok(())
}

async fn save_res(
    dates: HashMap<String, Date>,
    res: HashMap<String, HashMap<Date, Attendance>>,
) -> anyhow::Result<()> {
    let opt_file = AsyncFileDialog::new()
        .add_filter("excel", &["xls", "xlsx"])
        .set_title("保存")
        .save_file()
        .await;

    if let Some(file_handle) = opt_file {
        let mut book = umya_spreadsheet::new_file();
        let worksheet = book.get_sheet_mut(&0).map_err(|e| anyhow::anyhow!(e))?;

        worksheet.get_cell_mut((1, 1)).set_value_string("工号");
        worksheet.get_cell_mut((2, 1)).set_value_string("姓名");

        let mut dates: Vec<Date> = dates.into_values().collect();
        dates.sort();

        let format = format_description!("[month padding:none]月[day padding:none]日");
        for (i, date) in dates.iter().enumerate() {
            let date_string = date.format(&format).unwrap();
            worksheet
                .get_cell_mut((2 * i as u32 + 3, 1))
                .set_value_string(format!("{}个人投入度", date_string));

            worksheet
                .get_cell_mut((2 * i as u32 + 4, 1))
                .set_value_string(format!("{}考勤\n（正常/不正常（缺卡、补卡、虚拟打卡、非主责项目或城市打卡），不正常说明原因）", date_string));

            worksheet
                .get_column_dimension_mut(&string_from_column_index(&(2 * i as u32 + 4)))
                .set_width(15_f64);
        }

        let mut all: Vec<(&String, &HashMap<Date, Attendance>)> = res.iter().collect();
        all.sort_by_key(|a| a.0);
        for (j, (employee_id, employee_attendance)) in all.iter().enumerate() {
            worksheet
                .get_cell_mut((1, j as u32 + 2))
                .set_value_string(*employee_id);

            for (i, date) in dates.iter().enumerate() {
                if let Some(attendance) = employee_attendance.get(date) {
                    if i == 0 {
                        worksheet
                            .get_cell_mut((2, j as u32 + 2))
                            .set_value_string(&attendance.employee_name);
                    }

                    worksheet
                        .get_cell_mut((2 * i as u32 + 3, j as u32 + 2))
                        .set_value_number(attendance.work_hours / 8);

                    //原因
                    worksheet
                        .get_cell_mut((2 * i as u32 + 4, j as u32 + 2))
                        .set_value_string(
                            format!(
                                "{}\n{}\n{}",
                                attendance.enter_info,
                                attendance.leave_info,
                                attendance.abnormal_reason
                            )
                            .trim(),
                        );
                }
            }
        }

        let (col, row) = worksheet.get_highest_column_and_row();
        let mut style = Style::default();
        let alignment = style.get_alignment_mut();
        alignment.set_vertical(VerticalAlignmentValues::Center);
        alignment.set_horizontal(HorizontalAlignmentValues::Center);
        alignment.set_wrap_text(true);

        worksheet.set_style_by_range(
            &format!(
                "{}:{}",
                coordinate_from_index(&1, &1),
                coordinate_from_index(&col, &row)
            ),
            style,
        );

        umya_spreadsheet::writer::xlsx::write(&book, file_handle.path())?;
    }

    Ok(())
}
