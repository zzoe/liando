use rfd::AsyncFileDialog;
use serde::{Deserialize, Serialize};
use slint::{PlatformError, SharedString};
use time::{Duration, OffsetDateTime};

slint::include_modules!();

mod attendance;
mod meetings;

fn main() -> Result<(), PlatformError> {
    let ui = App::new()?;

    init_date(&ui);
    on_statistics_file_select(&ui);
    on_record_file_select(&ui);
    on_execute_clicked(&ui);

    ui.run()
}

fn init_date(ui: &App) {
    let today = OffsetDateTime::now_local().map(|d| d.date());
    let last_friday = today.map(|d| {
        d.saturating_sub(Duration::days(
            d.weekday().number_days_from_sunday() as i64 + 2_i64,
        ))
    });
    let last_monday = last_friday
        .map(|d| d.saturating_sub(Duration::days(4)).to_string().into())
        .unwrap_or_default();
    let last_friday = last_friday
        .map(|d| SharedString::from(d.to_string()))
        .unwrap_or_default();

    ui.set_start_date(last_monday);
    ui.set_end_date(last_friday);
}

fn on_statistics_file_select(ui: &App) {
    let ui_weak = ui.as_weak();
    ui.on_statistics_select_clicked(move || {
        let ui_weak_copy = ui_weak.clone();
        async_global_executor::spawn(select_file(
            ui_weak_copy,
            FileClassification::DailyStatistics,
        ))
        .detach();
    });
}

fn on_record_file_select(ui: &App) {
    let ui_weak = ui.as_weak();
    ui.on_record_select_clicked(move || {
        let ui_weak_copy = ui_weak.clone();
        async_global_executor::spawn(select_file(
            ui_weak_copy,
            FileClassification::OriginalRecord,
        ))
        .detach();
    });
}

fn on_execute_clicked(ui: &App) {
    let ui_weak = ui.as_weak();
    ui.on_execute_clicked(move || {
        let ui_weak_copy = ui_weak.clone();
        async_global_executor::spawn(async {
            meetings::query_tasks(ui_weak_copy).await;
        })
        .detach();
    });
}

enum FileClassification {
    DailyStatistics,
    OriginalRecord,
}

async fn select_file(ui_weak: slint::Weak<App>, file_classification: FileClassification) {
    let opt_file = AsyncFileDialog::new()
        .add_filter("excel", &["xls", "xlsx"])
        .set_title("请选择考勤Excel")
        .pick_file()
        .await;

    ui_weak
        .upgrade_in_event_loop(move |ui| {
            if let Some(file) = opt_file {
                let file_path = SharedString::from(file.path().to_str().unwrap_or_default());
                match file_classification {
                    FileClassification::DailyStatistics => ui.set_statistics_file(file_path),
                    FileClassification::OriginalRecord => ui.set_record_file(file_path),
                }
            }

            ui.set_button_enabled(true);
        })
        .unwrap();
}
