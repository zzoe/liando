use crate::App;
use serde::{Deserialize, Serialize};

pub(crate) async fn generate_report(ui_weak: slint::Weak<App>) {
    let (s, r) = async_channel::bounded(1);
    let ui_weak_copy = ui_weak.clone();
    async_global_executor::spawn_blocking(move || {
        ui_weak_copy
            .upgrade_in_event_loop(move |ui| {
                let req = ReportReq {
                    start_date: ui.get_start_date().to_string(),
                    end_date: ui.get_end_date().to_string(),
                    statistics_file: ui.get_statistics_file().to_string(),
                    record_file: ui.get_record_file().to_string(),
                };
                s.send_blocking(req).unwrap();
            })
            .unwrap()
    })
    .await;

    let report_req = r.recv().await.unwrap_or_default();
}

#[derive(Default, Deserialize, Serialize)]
struct ReportReq {
    start_date: String,
    end_date: String,
    statistics_file: String,
    record_file: String,
}
