#![windows_subsystem = "windows"]
use rfd::AsyncFileDialog;
use slint::{PhysicalPosition, PlatformError, SharedString, ComponentHandle};
use time::{Duration, OffsetDateTime};

slint::include_modules!();

mod attendance;

fn main() -> Result<(), PlatformError> {
    let app = App::new()?;

    init_date(&app);
    on_statistics_file_select(&app);
    on_record_file_select(&app);
    on_execute_clicked(&app);

    app.window().set_position(PhysicalPosition::new(520, 520));

    app.run()
}

fn init_date(app: &App) {
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

    app.global::<Logic>().set_start_date(last_monday);
    app.global::<Logic>().set_end_date(last_friday);
}

fn on_statistics_file_select(app: &App) {
    let app_weak = app.as_weak();
    app.global::<Logic>().on_statistics_import_clicked(move || {
        let app_weak_copy = app_weak.clone();
        async_global_executor::spawn(select_file(
            app_weak_copy,
            FileClassification::DailyStatistics,
        ))
        .detach();
    });
}

fn on_record_file_select(app: &App) {
    let app_weak = app.as_weak();
    app.global::<Logic>().on_record_import_clicked(move || {
        let app_weak_copy = app_weak.clone();
        async_global_executor::spawn(select_file(
            app_weak_copy,
            FileClassification::OriginalRecord,
        ))
        .detach();
    });
}

fn on_execute_clicked(app: &App) {
    let app_weak = app.as_weak();
    app.global::<Logic>().on_home_execute_clicked(move || {
        let app_weak_copy = app_weak.clone();
        async_global_executor::spawn(attendance::execute_handle(app_weak_copy)).detach();
    });
}

enum FileClassification {
    DailyStatistics,
    OriginalRecord,
}

async fn select_file(app_weak: slint::Weak<App>, file_classification: FileClassification) {
    let opt_file = AsyncFileDialog::new()
        .add_filter("excel", &["xls", "xlsx"])
        .set_title("请选择考勤Excel")
        .pick_file()
        .await;

    // 根据文件类型分别解析文件

    app_weak
        .upgrade_in_event_loop(move |app| {
            app.global::<Logic>().set_button_enabled(true);
        })
        .unwrap();
}
