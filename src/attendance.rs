use std::collections::HashMap;
use std::path::{Path, PathBuf};

use anyhow::{anyhow, Result};
use async_std::channel::Sender;
use async_std::{channel, task};
use rfd::AsyncFileDialog;
use sled::{open, Db};
use slint::{ComponentHandle, PhysicalPosition, SharedString};
use speedy::{Readable, Writable};
use time::{macros::format_description, Date, Duration};
use time::{OffsetDateTime, UtcOffset};
use umya_spreadsheet::helper::coordinate::{column_index_from_string, coordinate_from_index, string_from_column_index};
use umya_spreadsheet::{HorizontalAlignmentValues, Style, VerticalAlignmentValues};

use crate::{Logic, Ui};

macro_rules! parse_input_col {
    ($ui:ident, $get:ident, $set:ident) => {{
        let old_value = $ui.global::<Logic>().$get();
        let col_str = get_3_alpha(&old_value);
        if col_str.is_empty() {
            None
        } else {
            let int_value = column_index_from_string(col_str);
            let new_value = int_value.to_string().into();
            if old_value.ne(&new_value) {
                $ui.global::<Logic>().$set(new_value);
            }
            Some(int_value)
        }
    }};
}

macro_rules! parse_input_row {
    ($ui:ident, $get:ident, $set:ident) => {{
        let old_value = $ui.global::<Logic>().$get();
        let int_value = old_value.parse::<u32>().unwrap_or_default();
        if int_value < 1 {
            None
        } else {
            let new_value = int_value.to_string().into();
            if old_value.ne(&new_value) {
                $ui.global::<Logic>().$set(new_value);
            }
            Some(int_value)
        }
    }};
}

pub(crate) struct App {
    ui: Ui,
    db: Db,
}

impl App {
    pub(crate) fn new() -> Self {
        App {
            ui: Ui::new().unwrap(),
            db: open("./liando.db").unwrap(),
        }
    }

    pub fn run(&self) -> Result<()> {
        self.init()?;

        self.ui.window().set_position(PhysicalPosition::new(520, 520));
        self.ui.run()?;

        Ok(())
    }

    fn init(&self) -> Result<()> {
        self.init_input();
        self.on_statistics_file_select();
        self.on_record_file_select();
        self.on_template_file_select();
        self.on_execute_clicked();

        Ok(())
    }

    fn init_input(&self) {
        let user_input = self
            .db
            .get("user_input")
            .ok()
            .flatten()
            .and_then(|value| UserInput::read_from_buffer(&value).ok())
            .unwrap_or_default();

        self.ui
            .global::<Logic>()
            .set_start_date(long_date_string(user_input.start_date));
        self.ui
            .global::<Logic>()
            .set_end_date(long_date_string(user_input.end_date));
        self.ui
            .global::<Logic>()
            .set_statistics_employee_id_col(string_from_column_index(&user_input.statistics_employee_id_col).into());
        self.ui
            .global::<Logic>()
            .set_statistics_date_col(string_from_column_index(&user_input.statistics_date_col).into());
        self.ui
            .global::<Logic>()
            .set_statistics_enter_result_col(string_from_column_index(&user_input.statistics_enter_result_col).into());
        self.ui
            .global::<Logic>()
            .set_statistics_leave_result_col(string_from_column_index(&user_input.statistics_leave_result_col).into());
        self.ui
            .global::<Logic>()
            .set_statistics_work_minutes_col(string_from_column_index(&user_input.statistics_work_minutes_col).into());
        self.ui
            .global::<Logic>()
            .set_statistics_start_row(user_input.statistics_start_row.to_string().into());
        self.ui
            .global::<Logic>()
            .set_record_employee_id_col(string_from_column_index(&user_input.record_employee_id_col).into());
        self.ui
            .global::<Logic>()
            .set_record_date_col(string_from_column_index(&user_input.record_date_col).into());
        self.ui
            .global::<Logic>()
            .set_record_abnormal_reason_col(string_from_column_index(&user_input.record_abnormal_reason_col).into());
        self.ui
            .global::<Logic>()
            .set_record_start_row(user_input.record_start_row.to_string().into());
    }

    fn on_statistics_file_select(&self) {
        let ui_weak = self.ui.as_weak();
        let db = self.db.clone();
        self.ui.global::<Logic>().on_statistics_import_clicked(move || {
            let ui_weak_copy1 = ui_weak.clone();
            let ui_weak_copy2 = ui_weak.clone();
            let db = db.clone();
            task::spawn(async move {
                if let Some(user_input) = get_input(ui_weak_copy1).await {
                    if let Some(file) = select_file("请选择每日统计表").await {
                        // todo 保存输入到sled
                        // 导入每日统计表到sled
                        if let Err(e) = update_statistics(file, &user_input, &db) {
                            ui_weak_copy2
                                .upgrade_in_event_loop(move |ui| {
                                    ui.set_alert_text(SharedString::from(e.to_string()));
                                    ui.invoke_alert();
                                })
                                .unwrap();
                        }
                    }
                }
            });
        });
    }

    fn on_record_file_select(&self) {
        let ui_weak = self.ui.as_weak();
        self.ui.global::<Logic>().on_record_import_clicked(move || {
            let ui_weak_copy = ui_weak.clone();
            task::spawn(async move {
                if let Some(file) = select_file("请选择原始记录表").await {
                    //
                } else {
                    ui_weak_copy
                        .upgrade_in_event_loop(|ui| {
                            ui.set_alert_text("未选择文件".into());
                            ui.invoke_alert();
                        })
                        .unwrap();
                }
            });
        });
    }

    fn on_template_file_select(&self) {}

    fn on_execute_clicked(&self) {
        let ui_weak = self.ui.as_weak();
        self.ui.global::<Logic>().on_home_execute_clicked(move || {
            let ui_weak_copy = ui_weak.clone();
            task::spawn(execute_handle(ui_weak_copy));
        });
    }
}

#[derive(Debug, Readable, Writable, PartialEq)]
struct UserInput {
    start_date: i32,
    end_date: i32,
    statistics_employee_id_col: u32,
    statistics_date_col: u32,
    statistics_enter_result_col: u32,
    statistics_leave_result_col: u32,
    statistics_work_minutes_col: u32,
    statistics_start_row: u32,
    record_employee_id_col: u32,
    record_date_col: u32,
    record_abnormal_reason_col: u32,
    record_start_row: u32,
}

impl Default for UserInput {
    fn default() -> Self {
        let today = OffsetDateTime::now_utc().to_offset(UtcOffset::from_hms(8, 0, 0).unwrap());
        let last_friday =
            today.saturating_sub(Duration::days(today.weekday().number_days_from_sunday() as i64 + 2_i64));
        let last_monday = last_friday.saturating_sub(Duration::days(4));

        UserInput {
            start_date: last_monday.to_julian_day(),
            end_date: last_friday.to_julian_day(),
            statistics_employee_id_col: 4,
            statistics_date_col: 7,
            statistics_enter_result_col: 10,
            statistics_leave_result_col: 12,
            statistics_work_minutes_col: 20,
            statistics_start_row: 5,
            record_employee_id_col: 4,
            record_date_col: 7,
            record_abnormal_reason_col: 14,
            record_start_row: 4,
        }
    }
}

impl UserInput {
    fn get_dates(&self) -> HashMap<String, Date> {
        let mut dates = HashMap::new();
        let mut loop_date = Date::from_julian_day(self.start_date).unwrap();
        let end_date = Date::from_julian_day(self.end_date).unwrap();
        let format = format_description!("[year repr:last_two]-[month]-[day]");
        while loop_date <= end_date {
            dates.insert(loop_date.format(&format).unwrap(), loop_date);
            loop_date = loop_date.saturating_add(Duration::days(1));
        }

        dates
    }
}

fn long_date_string(date: i32) -> SharedString {
    let format = format_description!("[year]-[month]-[day]");
    SharedString::from(Date::from_julian_day(date).unwrap().format(&format).unwrap())
}

async fn select_file(title: &str) -> Option<PathBuf> {
    AsyncFileDialog::new()
        .add_filter("excel", &["xls", "xlsx"])
        .set_title(title)
        .pick_file()
        .await
        .map(|file| file.path().to_owned())
}

fn get_3_alpha(ss: &SharedString) -> String {
    ss.chars()
        .filter(|&c| c.is_ascii_alphabetic())
        .take(3)
        .collect::<String>()
}

async fn get_input(ui_weak: slint::Weak<Ui>) -> Option<UserInput> {
    let (s, r) = channel::bounded(1);
    ui_weak
        .upgrade_in_event_loop(move |ui| {
            if let Err(e) = parse_input(&ui, &s) {
                ui.set_alert_text(SharedString::from(e.to_string()));
                ui.invoke_alert();
            }
            s.close();
        })
        .unwrap();

    r.recv().await.ok()
}

fn parse_input(ui: &Ui, sender: &Sender<UserInput>) -> Result<()> {
    let format = format_description!("[year]-[month]-[day]");

    let mut start_date_str = ui.global::<Logic>().get_start_date();
    let opt_start_date = Date::parse(&start_date_str, &format).map_err(|_| anyhow!("开始日期，填写有误，请检查"))?;

    let mut end_date_str = ui.global::<Logic>().get_end_date();
    let opt_end_date = Date::parse(&end_date_str, &format).map_err(|_| anyhow!("结束日期，填写有误，请检查"))?;

    let statistics_employee_id_col =
        parse_input_col!(ui, get_statistics_employee_id_col, set_statistics_employee_id_col)
            .ok_or(anyhow!("每日统计表-工号，填写有误，请检查"))?;
    let statistics_date_col = parse_input_col!(ui, get_statistics_date_col, set_statistics_date_col)
        .ok_or(anyhow!("每日统计表-日期，填写有误，请检查"))?;
    let statistics_enter_result_col =
        parse_input_col!(ui, get_statistics_enter_result_col, set_statistics_enter_result_col)
            .ok_or(anyhow!("每日统计表-上班-打卡结果1，填写有误，请检查"))?;
    let statistics_leave_result_col =
        parse_input_col!(ui, get_statistics_leave_result_col, set_statistics_leave_result_col)
            .ok_or(anyhow!("每日统计表-下班-打卡结果1，填写有误，请检查"))?;
    let statistics_work_minutes_col =
        parse_input_col!(ui, get_statistics_work_minutes_col, set_statistics_work_minutes_col)
            .ok_or(anyhow!("每日统计表-工作时长(分钟)，填写有误，请检查"))?;
    let statistics_start_row = parse_input_row!(ui, get_statistics_start_row, set_statistics_start_row)
        .ok_or(anyhow!("每日统计表，数据起始行号，填写有误，请检查"))?;
    let record_employee_id_col = parse_input_col!(ui, get_record_employee_id_col, set_record_employee_id_col)
        .ok_or(anyhow!("原始记录表-工号，填写有误，请检查"))?;
    let record_date_col = parse_input_col!(ui, get_record_date_col, set_record_date_col)
        .ok_or(anyhow!("原始记录表-日期，填写有误，请检查"))?;
    let record_abnormal_reason_col =
        parse_input_col!(ui, get_record_abnormal_reason_col, set_record_abnormal_reason_col)
            .ok_or(anyhow!("原始记录表-异常打卡原因，填写有误，请检查"))?;
    let record_start_row = parse_input_row!(ui, get_record_start_row, set_record_start_row)
        .ok_or(anyhow!("原始记录表，数据起始行号，填写有误，请检查"))?;

    let mut user_input = UserInput {
        start_date: opt_start_date.to_julian_day(),
        end_date: opt_end_date.to_julian_day(),
        statistics_employee_id_col,
        statistics_date_col,
        statistics_enter_result_col,
        statistics_leave_result_col,
        statistics_work_minutes_col,
        statistics_start_row,
        record_employee_id_col,
        record_date_col,
        record_abnormal_reason_col,
        record_start_row,
    };

    if start_date_str > end_date_str {
        std::mem::swap(&mut start_date_str, &mut end_date_str);
        std::mem::swap(&mut user_input.start_date, &mut user_input.end_date);
        ui.global::<Logic>().set_start_date(start_date_str);
        ui.global::<Logic>().set_end_date(end_date_str);
    }

    sender.send_blocking(user_input)?;

    Ok(())
}

fn update_statistics(path: impl AsRef<Path>, user_input: &UserInput, db: &Db) -> Result<()> {
    let book = umya_spreadsheet::reader::xlsx::read(path.as_ref())?;
    let worksheet = book.get_sheet(&0).map_err(|e| anyhow!(e))?;
    let (_, max_row) = worksheet.get_highest_column_and_row();
    // let mut res = HashMap::new();
    let format = format_description!("[year]-[month]-[day]");

    for r in user_input.statistics_start_row..max_row + 1 {
        let employee_id = worksheet.get_formatted_value((user_input.statistics_employee_id_col, r));
        if employee_id.is_empty() {
            continue;
        }

        if let Some(attendance_date) = worksheet
            .get_formatted_value((user_input.statistics_date_col, r))
            .get(..8)
        {
            println!("attendance_date: {attendance_date}");
            let date = Date::parse(&format!("20{attendance_date}"), &format);
            println!("date: {date:?}");

            //     if !res.contains_key(&employee_id) {
            //         res.insert(employee_id.to_string(), HashMap::new());
            //     }
            //
            //     let employee_attendance = res.get_mut(&employee_id).unwrap();
            //     employee_attendance.insert(
            //         *date,
            //         Attendance {
            //             employee_id,
            //             enter_info: worksheet.get_formatted_value((user_input.statistics_enter_result_col, r)),
            //             leave_info: worksheet.get_formatted_value((user_input.statistics_leave_result_col, r)),
            //             abnormal_reason: String::new(),
            //             work_hours: worksheet.get_value_number((user_input.statistics_work_minutes_col, r)).unwrap_or_default(),
            //         },
            //     );
        }
    }

    Ok(())
}

pub(crate) async fn execute_handle(ui_weak: slint::Weak<Ui>) {
    let ui_weak_copy = ui_weak.clone();
    if let Some(user_input) = get_input(ui_weak).await {
        if let Err(e) = generate_report(user_input).await {
            ui_weak_copy
                .upgrade_in_event_loop(move |ui| {
                    ui.set_alert_text(e.to_string().into());
                    ui.invoke_alert();
                })
                .unwrap();
        }
    }

    ui_weak_copy
        .upgrade_in_event_loop(move |ui| {
            ui.global::<Logic>().set_button_enabled(true);
        })
        .unwrap();
}

async fn generate_report(user_input: UserInput) -> Result<()> {
    // let dates = user_input.get_dates();
    // //从每日统计文件取每人每天上下班的考勤状态(G,I)和工作时长（Q）
    // let mut all_attendance = get_everyone_statistics(Path::new(&user_input.statistics_file), &dates)?;
    // //从原始记录文件取每人每天上下班的虚拟打卡情况(N)
    // get_everyone_abnormal_reason(Path::new(&user_input.record_file), &dates, &mut all_attendance)?;
    // //输出文件
    // save_res(dates, all_attendance).await?;

    /*
       1. 默认Sheet取的第一页
       2. 当天上/下班既有正常打卡，也有虚拟打卡，是否记虚拟?
       3. 各种状态的映射关系?

       可配置项：每个值的列、打卡结果映射成枚举
    */
    Ok(())
}

// #[derive(Readable, Writable, PartialEq)]
struct Attendance {
    employee_id: String,
    enter_info: String,
    leave_info: String,
    work_hours: f64,
    abnormal_reason: String,
}

// HashMap<工号，HashMap<日期， 考勤情况>>
fn get_everyone_statistics(
    path: impl AsRef<Path>,
    dates: &HashMap<String, Date>,
) -> Result<HashMap<String, HashMap<Date, Attendance>>> {
    let book = umya_spreadsheet::reader::xlsx::read(path.as_ref())?;
    let worksheet = book.get_sheet(&0).map_err(|e| anyhow!(e))?;
    let (_, max_row) = worksheet.get_highest_column_and_row();
    let mut res = HashMap::new();

    for r in 5..max_row + 1 {
        if let Some(attendance_date) = worksheet.get_formatted_value((7, r)).get(..8) {
            // println!("attendance_date: {attendance_date}");
            if let Some(date) = dates.get(attendance_date) {
                let employee_id = worksheet.get_formatted_value((4, r));
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
                        // employee_name: worksheet.get_formatted_value((1, r)),
                        employee_id,
                        enter_info: worksheet.get_formatted_value((10, r)),
                        leave_info: worksheet.get_formatted_value((12, r)),
                        abnormal_reason: String::new(),
                        work_hours: worksheet.get_value_number((20, r)).unwrap_or_default(),
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
) -> Result<()> {
    let book = umya_spreadsheet::reader::xlsx::read(path.as_ref())?;
    let worksheet = book.get_sheet(&0).map_err(|e| anyhow!(e))?;
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

async fn save_res(dates: HashMap<String, Date>, res: HashMap<String, HashMap<Date, Attendance>>) -> Result<()> {
    let opt_file = AsyncFileDialog::new()
        .add_filter("excel", &["xls", "xlsx"])
        .set_title("保存")
        .save_file()
        .await;

    if let Some(file_handle) = opt_file {
        let mut book = umya_spreadsheet::new_file();
        let worksheet = book.get_sheet_mut(&0).map_err(|e| anyhow!(e))?;

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

            worksheet.get_cell_mut((2 * i as u32 + 4, 1)).set_value_string(format!(
                "{}考勤\n（正常/不正常（缺卡、补卡、虚拟打卡、非主责项目或城市打卡），不正常说明原因）",
                date_string
            ));

            worksheet
                .get_column_dimension_mut(&string_from_column_index(&(2 * i as u32 + 4)))
                .set_width(15_f64);
        }

        let mut all: Vec<(&String, &HashMap<Date, Attendance>)> = res.iter().collect();
        all.sort_by_key(|a| a.0);
        for (j, (employee_id, employee_attendance)) in all.iter().enumerate() {
            worksheet.get_cell_mut((1, j as u32 + 2)).set_value_string(*employee_id);

            for (i, date) in dates.iter().enumerate() {
                if let Some(attendance) = employee_attendance.get(date) {
                    // if i == 0 {
                    //     worksheet
                    //         .get_cell_mut((2, j as u32 + 2))
                    //         .set_value_string(&attendance.employee_name);
                    // }

                    worksheet
                        .get_cell_mut((2 * i as u32 + 3, j as u32 + 2))
                        .set_value_number(attendance.work_hours / 60.0);

                    //原因
                    worksheet
                        .get_cell_mut((2 * i as u32 + 4, j as u32 + 2))
                        .set_value_string(
                            format!(
                                "{}\n{}\n{}",
                                sum_up(&attendance.enter_info)
                                    .replace("缺卡", "缺早卡")
                                    .replace("补卡", "补早卡"),
                                sum_up(&attendance.leave_info)
                                    .replace("缺卡", "缺晚卡")
                                    .replace("补卡", "补晚卡"),
                                sum_up(&attendance.abnormal_reason)
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

fn sum_up(reason: &str) -> String {
    for i in ["缺卡", "补卡", "迟到", "早退", "虚拟"] {
        if reason.contains(i) {
            return i.to_string();
        }
    }
    String::new()
}
