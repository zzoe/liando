use crate::App;
use serde::{Deserialize, Serialize};
use slint::SharedString;

pub(crate) async fn query_tasks(ui_weak: slint::Weak<App>) {
    let (s, r) = async_channel::bounded(1);
    let ui_weak_copy = ui_weak.clone();
    async_global_executor::spawn_blocking(move || {
        ui_weak_copy
            .upgrade_in_event_loop(move |ui| {
                let req = TaskReq {
                    page_num: 1,
                    page_size: 2000,
                    search_type: 4,
                    start_date: ui.get_start_date().to_string(),
                    end_date: ui.get_end_date().to_string(),
                    authorization: ui.get_authorization().to_string(),
                };
                s.send_blocking(req).unwrap();
            })
            .unwrap()
    })
    .await;

    let task_req = r.recv().await.unwrap_or_default();

    //通过surf client 发送http请求

    let task_res = surf::get("http://tmp.liando.cn/api/inner/business/tsTask/list")
        .query(&task_req)
        .unwrap()
        .header("Authorization", task_req.authorization)
        .recv_string()
        .await;

    ui_weak
        .upgrade_in_event_loop(|ui| {
            if let Ok(tasks) = task_res {
                ui.set_tasks(SharedString::from(tasks));
            }
            ui.set_button_enabled(true);
        })
        .unwrap();
}

#[derive(Default, Deserialize, Serialize)]
struct TaskReq {
    #[serde(rename = "pageNum")]
    page_num: u32,
    #[serde(rename = "pageSize")]
    page_size: u32,
    #[serde(rename = "searchType")]
    search_type: u32,
    #[serde(rename = "startTime")]
    start_date: String,
    #[serde(rename = "endTime")]
    end_date: String,
    #[serde(skip)]
    authorization: String,
}

// struct Response<T> {
//     code: String,
//     data: T,
// }
//
// struct TaskRes {}
