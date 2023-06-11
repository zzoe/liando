#![windows_subsystem = "windows"]
use rfd::AsyncFileDialog;
use slint::{PhysicalPosition, PlatformError, SharedString};
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
    // use i_slint_backend_winit::WinitWindowAccessor;
    // app.show()?;
    //
    // if let Some(s) = app
    //     .window()
    //     .with_winit_window(|w| w.primary_monitor().map(|h| h.size()).unwrap_or_default())
    // {
    //     app.window().set_position(PhysicalPosition::new(
    //         ((s.width - 800) / 2) as i32,
    //         ((s.height - 240) / 2) as i32,
    //     ));
    // }
    // println!("position: {:#?}", app.window().position());
    //
    // slint::run_event_loop()?;
    // app.hide()
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

    app.set_start_date(last_monday);
    app.set_end_date(last_friday);
}

fn on_statistics_file_select(app: &App) {
    let app_weak = app.as_weak();
    app.on_statistics_select_clicked(move || {
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
    app.on_record_select_clicked(move || {
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
    app.on_execute_clicked(move || {
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

    app_weak
        .upgrade_in_event_loop(move |app| {
            if let Some(file) = opt_file {
                let file_path = SharedString::from(file.path().to_str().unwrap_or_default());
                match file_classification {
                    FileClassification::DailyStatistics => app.set_statistics_file(file_path),
                    FileClassification::OriginalRecord => app.set_record_file(file_path),
                }
            }

            app.set_button_enabled(true);
        })
        .unwrap();
}
