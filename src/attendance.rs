use std::collections::HashMap;
use std::path::{Path, PathBuf};

use anyhow::{anyhow, Result};
use async_std::channel::Sender;
use async_std::{channel, task};
use rfd::AsyncFileDialog;
use sled::{open, Db};
use slint::{ComponentHandle, Model, ModelRc, PhysicalPosition, SharedString, VecModel};
use speedy::{Readable, Writable};
use time::{macros::format_description, Date, Duration};
use time::{OffsetDateTime, UtcOffset};
use umya_spreadsheet::helper::coordinate::{
    column_index_from_string, coordinate_from_index, string_from_column_index,
};
use umya_spreadsheet::{HorizontalAlignmentValues, Style, VerticalAlignmentValues};

use crate::{Logic, TemplateConfig, Ui};

macro_rules! parse_input_col {
    ($ui:ident, $get:ident, $set:ident) => {{
        let old_value = $ui.global::<Logic>().$get();
        let col_str = get_3_alpha(&old_value);
        if col_str.is_empty() {
            None
        } else {
            let int_value = column_index_from_string(col_str);
            let new_value = string_from_column_index(&int_value).into();
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

        self.ui
            .window()
            .set_position(PhysicalPosition::new(520, 520));
        self.ui.run()?;

        Ok(())
    }

    fn init(&self) -> Result<()> {
        self.init_input();
        self.on_statistics_file_select();
        self.on_record_file_select();
        self.on_template_remove_clicked();
        self.on_template_push_clicked();
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
        self.ui.global::<Logic>().set_statistics_employee_id_col(
            string_from_column_index(&user_input.statistics_employee_id_col).into(),
        );
        self.ui.global::<Logic>().set_statistics_date_col(
            string_from_column_index(&user_input.statistics_date_col).into(),
        );
        self.ui.global::<Logic>().set_statistics_enter_result_col(
            string_from_column_index(&user_input.statistics_enter_result_col).into(),
        );
        self.ui.global::<Logic>().set_statistics_leave_result_col(
            string_from_column_index(&user_input.statistics_leave_result_col).into(),
        );
        self.ui.global::<Logic>().set_statistics_work_minutes_col(
            string_from_column_index(&user_input.statistics_work_minutes_col).into(),
        );
        self.ui
            .global::<Logic>()
            .set_statistics_start_row(user_input.statistics_start_row.to_string().into());
        self.ui.global::<Logic>().set_record_employee_id_col(
            string_from_column_index(&user_input.record_employee_id_col).into(),
        );
        self.ui
            .global::<Logic>()
            .set_record_date_col(string_from_column_index(&user_input.record_date_col).into());
        self.ui.global::<Logic>().set_record_abnormal_reason_col(
            string_from_column_index(&user_input.record_abnormal_reason_col).into(),
        );
        self.ui
            .global::<Logic>()
            .set_record_start_row(user_input.record_start_row.to_string().into());

        let template_cfg = user_input
            .template_cfg
            .iter()
            .fold(Vec::new(), |mut cfg, value| {
                cfg.push(TemplateConfig {
                    template_employee_id_col: string_from_column_index(&value.0).into(),
                    template_start_col: string_from_column_index(&value.1).into(),
                    template_title_row: value.2.to_string().into(),
                });
                cfg
            });
        self.ui
            .global::<Logic>()
            .set_template_configs(ModelRc::new(VecModel::from(template_cfg)));
    }

    fn on_statistics_file_select(&self) {
        let ui_weak = self.ui.as_weak();
        let db = self.db.clone();
        self.ui
            .global::<Logic>()
            .on_statistics_import_clicked(move || {
                let ui_weak_copy1 = ui_weak.clone();
                let ui_weak_copy2 = ui_weak.clone();
                let db = db.clone();
                task::spawn(async move {
                    if let Some(user_input) = get_input(ui_weak_copy1).await {
                        if let Some(file) = select_file("请选择每日统计表").await {
                            // 保存输入，并导入上下班情况和工作时长到sled
                            let res = update_statistics(file, &user_input, &db);
                            reset_button(ui_weak_copy2, res);
                        }
                    }
                });
            });
    }

    fn on_record_file_select(&self) {
        let ui_weak = self.ui.as_weak();
        let db = self.db.clone();
        self.ui.global::<Logic>().on_record_import_clicked(move || {
            let ui_weak1 = ui_weak.clone();
            let ui_weak2 = ui_weak.clone();
            let db = db.clone();
            task::spawn(async move {
                if let Some(user_input) = get_input(ui_weak1).await {
                    if let Some(file) = select_file("请选择原始记录表").await {
                        // 保存输入，并导入考勤异常原因到sled
                        let res = update_record(file, &user_input, &db);
                        reset_button(ui_weak2, res);
                    }
                }
            });
        });
    }

    fn on_template_remove_clicked(&self) {
        let ui_weak = self.ui.as_weak();
        self.ui
            .global::<Logic>()
            .on_template_remove_clicked(move |index| {
                ui_weak
                    .upgrade_in_event_loop(move |ui| {
                        let mut template_cfg = ui
                            .global::<Logic>()
                            .get_template_configs()
                            .iter()
                            .collect::<Vec<TemplateConfig>>();
                        template_cfg.remove(index as usize);

                        ui.global::<Logic>()
                            .set_template_configs(ModelRc::new(VecModel::from(template_cfg)));
                        ui.global::<Logic>().set_button_enabled(true);
                    })
                    .ok();
            });
    }

    fn on_template_push_clicked(&self) {
        let ui_weak = self.ui.as_weak();
        self.ui.global::<Logic>().on_template_push_clicked(move || {
            ui_weak
                .upgrade_in_event_loop(move |ui| {
                    let mut template_cfg = ui
                        .global::<Logic>()
                        .get_template_configs()
                        .iter()
                        .collect::<Vec<TemplateConfig>>();
                    template_cfg.push(TemplateConfig::default());

                    ui.global::<Logic>()
                        .set_template_configs(ModelRc::new(VecModel::from(template_cfg)));
                    ui.global::<Logic>().set_button_enabled(true);
                })
                .ok();
        });
    }

    fn on_execute_clicked(&self) {
        let ui_weak = self.ui.as_weak();
        let db = self.db.clone();
        self.ui.global::<Logic>().on_home_execute_clicked(move || {
            let ui_weak1 = ui_weak.clone();
            let ui_weak2 = ui_weak.clone();
            let db = db.clone();
            task::spawn(async move {
                if let Some(user_input) = get_input(ui_weak1).await {
                    if let Some(file) = select_file("请选择今日份模板作为导出文件").await
                    {
                        // 保存输入，并根据sled信息生成结果
                        let res = generate_report(file, &user_input, &db);
                        reset_button(ui_weak2, res);
                    }
                }
            });
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
    template_cfg: Vec<(u32, u32, u32)>,
}

impl Default for UserInput {
    fn default() -> Self {
        let today = OffsetDateTime::now_utc().to_offset(UtcOffset::from_hms(8, 0, 0).unwrap());
        let last_friday = today.saturating_sub(Duration::days(
            today.weekday().number_days_from_sunday() as i64 + 2_i64,
        ));
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
            template_cfg: vec![(1, 7, 2), (4, 9, 2), (4, 9, 2)],
        }
    }
}

#[derive(Debug, Default, Readable, Writable, PartialEq)]
struct Attendance {
    employee_id: String,
    enter_info: String,
    leave_info: String,
    work_minutes: f64,
    abnormal_reason: String,
}

fn long_date_string(date: i32) -> SharedString {
    let format = format_description!("[year]-[month]-[day]");
    SharedString::from(
        Date::from_julian_day(date)
            .unwrap()
            .format(&format)
            .unwrap(),
    )
}

fn get_3_alpha(ss: &SharedString) -> String {
    ss.chars()
        .filter(|&c| c.is_ascii_alphabetic())
        .take(3)
        .map(|c| c.to_ascii_uppercase())
        .collect::<String>()
}

async fn select_file(title: &str) -> Option<PathBuf> {
    AsyncFileDialog::new()
        .add_filter("excel", &["xls", "xlsx"])
        .set_title(title)
        .pick_file()
        .await
        .map(|file| file.path().to_owned())
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

fn parse_col(old_value: SharedString) -> Option<u32> {
    let col_str = get_3_alpha(&old_value);
    if col_str.is_empty() {
        None
    } else {
        // let int_value = column_index_from_string(col_str);
        // let new_value: SharedString = int_value.to_string().into();
        Some(column_index_from_string(col_str))
    }
}

fn parse_input(ui: &Ui, sender: &Sender<UserInput>) -> Result<()> {
    let format = format_description!("[year]-[month]-[day]");

    let mut start_date_str = ui.global::<Logic>().get_start_date();
    let opt_start_date =
        Date::parse(&start_date_str, &format).map_err(|_| anyhow!("开始日期，填写有误，请检查"))?;

    let mut end_date_str = ui.global::<Logic>().get_end_date();
    let opt_end_date =
        Date::parse(&end_date_str, &format).map_err(|_| anyhow!("结束日期，填写有误，请检查"))?;

    let statistics_employee_id_col = parse_input_col!(
        ui,
        get_statistics_employee_id_col,
        set_statistics_employee_id_col
    )
    .ok_or(anyhow!("每日统计表-工号，填写有误，请检查"))?;
    let statistics_date_col =
        parse_input_col!(ui, get_statistics_date_col, set_statistics_date_col)
            .ok_or(anyhow!("每日统计表-日期，填写有误，请检查"))?;
    let statistics_enter_result_col = parse_input_col!(
        ui,
        get_statistics_enter_result_col,
        set_statistics_enter_result_col
    )
    .ok_or(anyhow!("每日统计表-上班-打卡结果1，填写有误，请检查"))?;
    let statistics_leave_result_col = parse_input_col!(
        ui,
        get_statistics_leave_result_col,
        set_statistics_leave_result_col
    )
    .ok_or(anyhow!("每日统计表-下班-打卡结果1，填写有误，请检查"))?;
    let statistics_work_minutes_col = parse_input_col!(
        ui,
        get_statistics_work_minutes_col,
        set_statistics_work_minutes_col
    )
    .ok_or(anyhow!("每日统计表-工作时长(分钟)，填写有误，请检查"))?;
    let statistics_start_row =
        parse_input_row!(ui, get_statistics_start_row, set_statistics_start_row)
            .ok_or(anyhow!("每日统计表，数据起始行号，填写有误，请检查"))?;
    let record_employee_id_col =
        parse_input_col!(ui, get_record_employee_id_col, set_record_employee_id_col)
            .ok_or(anyhow!("原始记录表-工号，填写有误，请检查"))?;
    let record_date_col = parse_input_col!(ui, get_record_date_col, set_record_date_col)
        .ok_or(anyhow!("原始记录表-日期，填写有误，请检查"))?;
    let record_abnormal_reason_col = parse_input_col!(
        ui,
        get_record_abnormal_reason_col,
        set_record_abnormal_reason_col
    )
    .ok_or(anyhow!("原始记录表-异常打卡原因，填写有误，请检查"))?;
    let record_start_row = parse_input_row!(ui, get_record_start_row, set_record_start_row)
        .ok_or(anyhow!("原始记录表，数据起始行号，填写有误，请检查"))?;

    let mut changed = false;
    let (template_cfg, template_cfg_str) = ui
        .global::<Logic>()
        .get_template_configs()
        .iter()
        .enumerate()
        .fold(
            (Vec::new(), Vec::new()),
            |(mut cfg, mut cfg_str), (i, v)| {
                let default_employee_id_col = if i == 0 { 1 } else { 4 };
                let default_start_col = if i == 0 { 7 } else { 9 };
                let default_title_row = 2;

                let old_value = v.clone();
                let parsed_value = (
                    parse_col(v.template_employee_id_col).unwrap_or(default_employee_id_col),
                    parse_col(v.template_start_col).unwrap_or(default_start_col),
                    v.template_title_row
                        .parse::<u32>()
                        .unwrap_or(default_title_row),
                );
                let parsed_value_str = TemplateConfig {
                    template_employee_id_col: string_from_column_index(&parsed_value.0).into(),
                    template_start_col: string_from_column_index(&parsed_value.1).into(),
                    template_title_row: parsed_value.2.to_string().into(),
                };

                changed |= old_value.ne(&parsed_value_str);

                cfg.push(parsed_value);
                cfg_str.push(parsed_value_str);

                (cfg, cfg_str)
            },
        );

    if changed {
        ui.global::<Logic>()
            .set_template_configs(ModelRc::new(VecModel::from(template_cfg_str)));
    }

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
        template_cfg,
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
    db.insert("user_input", user_input.write_to_vec()?)?;
    let format = format_description!("[year]-[month]-[day]");
    let book = umya_spreadsheet::reader::xlsx::read(path.as_ref())?;
    let worksheet = book.get_sheet(&0).map_err(|e| anyhow!(e))?;
    let (_, max_row) = worksheet.get_highest_column_and_row();

    for r in user_input.statistics_start_row..max_row + 1 {
        let employee_id = worksheet.get_formatted_value((user_input.statistics_employee_id_col, r));
        if employee_id.is_empty() {
            continue;
        }

        if let Some(attendance_date) = worksheet
            .get_formatted_value((user_input.statistics_date_col, r))
            .get(..8)
            .and_then(|date| Date::parse(&format!("20{date}"), &format).ok())
        {
            // println!("{attendance_date}_{employee_id}");
            db.fetch_and_update(&format!("{attendance_date}_{employee_id}"), |old| {
                let mut attendance = old
                    .and_then(|value| Attendance::read_from_buffer(value).ok())
                    .unwrap_or_default();
                attendance.employee_id = employee_id.clone();
                attendance.enter_info =
                    worksheet.get_formatted_value((user_input.statistics_enter_result_col, r));
                attendance.leave_info =
                    worksheet.get_formatted_value((user_input.statistics_leave_result_col, r));
                attendance.work_minutes = worksheet
                    .get_value_number((user_input.statistics_work_minutes_col, r))
                    .unwrap_or_default();
                attendance.write_to_vec().ok()
            })?;
        }
    }

    Ok(())
}

fn update_record(path: impl AsRef<Path>, user_input: &UserInput, db: &Db) -> Result<()> {
    db.insert("user_input", user_input.write_to_vec()?)?;
    let format = format_description!("[year]-[month]-[day]");
    let book = umya_spreadsheet::reader::xlsx::read(path.as_ref())?;
    let worksheet = book.get_sheet(&0).map_err(|e| anyhow!(e))?;
    let (_, max_row) = worksheet.get_highest_column_and_row();

    for r in user_input.record_start_row..max_row + 1 {
        let employee_id = worksheet.get_formatted_value((user_input.record_employee_id_col, r));
        if employee_id.is_empty() {
            continue;
        }

        if let Some(attendance_date) = worksheet
            .get_formatted_value((user_input.record_date_col, r))
            .get(..8)
            .and_then(|date| Date::parse(&format!("20{date}"), &format).ok())
        {
            db.fetch_and_update(&format!("{attendance_date}_{employee_id}"), |old| {
                let mut attendance = old
                    .and_then(|value| Attendance::read_from_buffer(value).ok())
                    .unwrap_or_default();
                attendance.employee_id = employee_id.clone();
                attendance.abnormal_reason =
                    worksheet.get_formatted_value((user_input.record_abnormal_reason_col, r));
                attendance.write_to_vec().ok()
            })?;
        }
    }

    Ok(())
}

fn generate_report(path: impl AsRef<Path>, user_input: &UserInput, db: &Db) -> Result<()> {
    db.insert("user_input", user_input.write_to_vec()?)?;
    let mut book = umya_spreadsheet::reader::xlsx::read(path.as_ref())?;

    for (sheet_index, template_cfg) in user_input.template_cfg.iter().enumerate() {
        let worksheet = book.get_sheet_mut(&sheet_index).map_err(|e| anyhow!(e))?;
        let (_, max_row) = worksheet.get_highest_column_and_row();
        let format = format_description!("[month padding:none]月[day padding:none]日");

        let mut loop_date = Date::from_julian_day(user_input.start_date).unwrap();
        let end_date = Date::from_julian_day(user_input.end_date).unwrap();
        let mut date_col = template_cfg.1;
        while loop_date <= end_date {
            let date_string = loop_date.format(&format).unwrap();
            let every_atd =
                db.scan_prefix(&format!("{loop_date}_"))
                    .fold(HashMap::new(), |mut map, kv| {
                        if let Ok((_, value)) = kv {
                            if let Ok(attendance) = Attendance::read_from_buffer(&value) {
                                map.insert(attendance.employee_id.clone(), attendance);
                            }
                        }
                        map
                    });

            for r in template_cfg.2..max_row + 1 {
                if r == template_cfg.2 {
                    // 写表头
                    worksheet
                        .get_cell_mut((date_col, r))
                        .set_value_string(format!("{}个人投入度", date_string));

                    worksheet.get_cell_mut((date_col + 1, r)).set_value_string(format!(
                        "{}考勤\n（正常/不正常（缺卡、补卡、虚拟打卡、非主责项目或城市打卡），不正常说明原因）",
                        date_string
                    ));

                    worksheet
                        .get_column_dimension_mut(&string_from_column_index(&(date_col + 1)))
                        .set_width(15_f64);

                    continue;
                }

                let employee_id = worksheet.get_formatted_value((template_cfg.0, r));
                if employee_id.is_empty() {
                    continue;
                }
                if let Some(attendance) = every_atd.get(&employee_id) {
                    worksheet
                        .get_cell_mut((date_col, r))
                        .set_value_number(attendance.work_minutes / 60.0);
                    worksheet.get_cell_mut((date_col + 1, r)).set_value_string(
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
                } else {
                    println!("无此人{employee_id}考勤信息");
                }
            }

            date_col += 2;
            loop_date = loop_date.saturating_add(Duration::days(1));
        }

        // 居中、换行
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
    }

    umya_spreadsheet::writer::xlsx::write(&book, path)?;
    Ok(())
}

fn reset_button(ui_weak: slint::Weak<Ui>, res: Result<()>) {
    ui_weak
        .upgrade_in_event_loop(move |ui| {
            if let Err(e) = res {
                ui.set_alert_text(SharedString::from(e.to_string()));
                ui.invoke_alert();
            }
            ui.global::<Logic>().set_button_enabled(true);
        })
        .ok();
}

fn sum_up(reason: &str) -> String {
    for i in ["缺卡", "补卡", "迟到", "早退", "虚拟"] {
        if reason.contains(i) {
            return i.to_string();
        }
    }
    String::new()
}
