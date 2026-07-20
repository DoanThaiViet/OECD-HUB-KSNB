/* ============================================================================
   OECD GOVERNANCE COMMAND CENTER — LỚP DỮ LIỆU CHUNG (single source of truth)
   ----------------------------------------------------------------------------
   - PH2 (Execution) GHI vào TASKS (biến JS, prototype — bản thật thay bằng
     SharePoint List); PH3 (Dashboard) & PH5 (Assessment) CHỈ ĐỌC-TÍNH.
   - Quan hệ 3 chiều: CRITERIA ↔ TASKS ↔ EVIDENCE (liên kết bằng mã).
   - Cập nhật dữ liệu: sửa các mảng bên dưới. KHÔNG sửa logic tính trong UI.
   ============================================================================ */
window.OECD = (function () {
  "use strict";

  /* ---------- Meta & SharePoint ---------- */
  const META = {
    updated: "17/07/2026",
    tracker: "Tracker KH 1017 ngày 10/07/2026",
    acgs: "Bộ ACGS Hybrid — CV 1152/TDKT-CL&QTRR ngày 20/05/2026",
    org: "Ban Kiểm soát Nội bộ · PVEP",
  };
  const SP_BASE = "https://pvepcorp.sharepoint.com/sites/m365-ksnb/Shared Documents/";
  const OECD_DIR = "8.5_OECD/";   // kho OECD trên M365-KSNB (đã dời từ 08_CƠ SỞ DỮ LIỆU CHUNG/... sang 8.5_OECD — xác minh 17/07/2026)
  const sp  = (rel) => encodeURI(SP_BASE + rel);
  const spO = (rel) => sp(OECD_DIR + rel);

  /* ---------- Trục thời gian KH 1017: idx 1 = T7/2026 … 10 = T4/2027 ---------- */
  const GMONTHS = ["T7/26","T8","T9","T10","T11","T12","T1/27","T2","T3","T4/27"];
  const NOW_IDX = 1; // T7/2026 — chỉnh khi sang tháng (2 = T8/2026 …)

  const GROUPS = {
    I:"Nâng cao tính chuyên nghiệp tiệm cận chuẩn mực OECD",
    II:"Hoàn thiện khung quản trị của PVEP",
    III:"Thực thi chiến lược, kế hoạch 5 năm & KH hằng năm",
    IV:"Quản trị rủi ro", V:"Kiểm tra giám sát & kiểm toán nội bộ",
    VI:"Phát triển bền vững (ESG)", VII:"Minh bạch công bố thông tin",
    VIII:"Quản trị Người đại diện",
  };

  /* ---------- TASKS: 21 nhiệm vụ KH QĐ 1017 (Phụ lục QĐ + tracker 10/07/2026) ----------
     PH2 ghi: p (tiến độ %), issues, requests, history[] — prototype lưu biến JS.        */
  const TASKS = [
    {id:"1.1", g:"I", t:"Tổ chức đào tạo quản trị công ty OECD", out:"Chương trình đào tạo theo lộ trình cho Lãnh đạo, quản lý cấp trung và đầu mối OECD các Ban/Đơn vị", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"QTNNL", co:"KSNB", dl:"Từ tháng 7/2026 (liên tục)", s:1, e:10, cont:true, p:31.7,
     prereq:"", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"Khởi tạo theo tracker: 31,7% — đã rà soát yêu cầu, đang xây dựng chương trình."}],
     criteria:["E1"]},
    {id:"1.2", g:"I", t:"Rà soát cơ cấu tổ chức, mô hình ủy ban trực thuộc HĐTV", out:"Báo cáo nghiên cứu, đánh giá cơ cấu tổ chức và đề xuất mô hình phù hợp", lead:"CT HĐTV Trương Quốc Lâm · TGĐ Nguyễn Thiện Bảo", own:"QTNNL", co:"CL&QTRR · KSNB", dl:"Quý III/2026", s:1, e:3, p:0,
     prereq:"Bám sát mô hình ủy ban đang triển khai tại Petrovietnam (NQ 5432 · PL1 mục 1.2)", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"Khởi tạo theo tracker: chưa bắt đầu."}],
     criteria:["E1"]},
    {id:"1.3", g:"I", t:"Hệ thống theo dõi thực thi nghị quyết, chỉ đạo Lãnh đạo PVEP", out:"Dashboard theo dõi NQ/QĐ/chỉ đạo + báo cáo trạng thái hằng tháng", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"KSNB (HĐTV) · CL&QTRR (BTGĐ)", co:"CN-CĐS · các Ban/VP", dl:"Quý III/2026", s:1, e:3, p:0,
     prereq:"", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"KSNB đã chuẩn hóa data dictionary (65%) và prototype dashboard (35%) ở lớp đầu việc."}],
     criteria:["E2"]},
    {id:"2.1", g:"II", milestone:true, t:"Rà soát hệ thống VBQLNB hiện hành", out:"Báo cáo đánh giá, đề xuất văn bản cần sửa đổi; KH hoàn thiện VBQLNB 2026 điều chỉnh", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"CL&QTRR", co:"Các Ban/VP", dl:"Tháng 8/2026", s:1, e:2, p:0,
     prereq:"Danh mục nền 79 VBQLNB theo CV 140/CVNB-HNT (30/04/2026)", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"Khởi tạo theo tracker: đang rà soát yêu cầu."}],
     criteria:["E1","D1"]},
    {id:"2.2", g:"II", t:"Đánh giá mức độ đáp ứng theo bộ tiêu chí Petrovietnam", out:"Báo cáo đánh giá mức độ đáp ứng", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"CL&QTRR", co:"Các Ban/VP", dl:"15 ngày sau khi PVN ban hành bộ tiêu chí", s:1, e:3, cont:true, p:0,
     prereq:"Petrovietnam ban hành bộ tiêu chí quản trị doanh nghiệp theo OECD (NQ 5432 · mục 2.5, từ quý III/2026)",
     issues:"Phụ thuộc tiến độ ban hành bộ tiêu chí của Petrovietnam.",
     requests:"Đề nghị đầu mối theo sát Ban KSNB Petrovietnam để nhận bộ tiêu chí sớm.",
     history:[{d:"10/07/2026", note:"Chưa bắt đầu — chờ bộ tiêu chí của Petrovietnam."}],
     criteria:[]},
    {id:"2.3", g:"II", t:"Đánh giá nội bộ mức độ trưởng thành quản trị hằng năm", out:"Báo cáo đánh giá nội bộ mức độ trưởng thành, đề xuất hoàn thiện", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"CL&QTRR", co:"QTNNL · KSNB", dl:"Tháng 1/2027", s:4, e:7, p:0,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"Chưa đến kỳ triển khai."}], criteria:[]},
    {id:"2.4", g:"II", t:"Chuẩn bị đánh giá độc lập mức độ trưởng thành", out:"Hoàn thiện hồ sơ phục vụ đánh giá độc lập", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"CL&QTRR", co:"QTNNL · KSNB", dl:"Quý I/2027", s:6, e:9, p:0,
     prereq:"Đánh giá độc lập toàn Tập đoàn dự kiến quý II/2027 (NQ 5432 · mục 2.3)", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"Chưa đến kỳ triển khai."}], criteria:[]},
    {id:"3.1", g:"III", t:"Kế hoạch thực thi Chiến lược / kế hoạch 5 năm 2026–2030", out:"Kế hoạch thực thi Chiến lược/kế hoạch 5 năm 2026–2030", lead:"TV HĐTV-TGĐ Nguyễn Thiện Bảo · PTGĐ Hoàng Xuân Dương", own:"CL&QTRR", co:"Các Ban/VP, Đơn vị", dl:"Quý III/2026", s:1, e:3, p:0,
     prereq:"Petrovietnam phê duyệt Chiến lược phát triển/kế hoạch 5 năm của PVEP", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"Khởi tạo theo tracker: chưa bắt đầu."}], criteria:[]},
    {id:"3.2", g:"III", t:"Cơ chế giám sát mục tiêu chiến lược, kế hoạch", out:"Bộ chỉ tiêu giám sát kèm ngưỡng cảnh báo sớm + dashboard quản trị cập nhật hằng tháng", lead:"TV HĐTV-TGĐ Nguyễn Thiện Bảo · PTGĐ Hoàng Xuân Dương", own:"CL&QTRR", co:"CN&CĐS · các Ban/VP, Đơn vị", dl:"Quý IV/2026", s:2, e:6, p:0,
     prereq:"Sau nhiệm vụ 3.1", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"Chưa bắt đầu."}], criteria:["E2"]},
    {id:"4.1", g:"IV", t:"Hoàn thiện hệ thống QTRR bao quát rủi ro trọng yếu", out:"Quy chế QTRR sửa đổi; danh mục rủi ro trọng yếu cấp PVEP; bộ KRI + ngưỡng cảnh báo sớm", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"CL&QTRR", co:"Các Ban/VP", dl:"Quý IV/2026", s:1, e:6, p:46.7,
     prereq:"", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"46,7% — đã rà soát (70%), đang dự thảo quy chế sửa đổi (50%)."}],
     criteria:["E4"]},
    {id:"4.2", g:"IV", t:"Nâng cao chất lượng báo cáo & hồ sơ rủi ro PVEP", out:"Báo cáo và hồ sơ rủi ro chuẩn hóa, bao quát đầy đủ rủi ro trọng yếu", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"CL&QTRR", co:"Các Đơn vị, Ban/VP (chủ sở hữu rủi ro)", dl:"Quý III/2026 · cập nhật định kỳ", s:1, e:3, cont:true, p:21.7,
     prereq:"", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"21,7% — đang chuẩn hóa biểu mẫu hồ sơ rủi ro."}], criteria:["E4"]},
    {id:"4.3", g:"IV", t:"Quản trị rủi ro số, an ninh mạng, SCADA ở cấp HĐTV", out:"Quy định CNTT/an toàn mạng sửa đổi; hồ sơ rủi ro tích hợp rủi ro số", lead:"TV HĐTV Đinh Trọng Huy · PTGĐ Ngô Khánh Xạ", own:"CN-CĐS", co:"CL&QTRR", dl:"Quý III/2026", s:1, e:3, p:0,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"Chưa bắt đầu."}], criteria:["E4"]},
    {id:"5.1", g:"V", t:"Chuẩn hóa quy trình KTGS & KTNB", out:"Quy trình KTGS, KTNB chuẩn hóa và số hóa theo chuẩn mực quốc tế (IIA)", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"KSNB", co:"—", dl:"Quý III/2026", s:1, e:3, p:0,
     prereq:"", issues:"", requests:"",
     history:[{d:"10/07/2026", note:"KSNB đã benchmark IIA và dự thảo biểu mẫu (20%) ở lớp đầu việc."}],
     criteria:["E3"]},
    {id:"5.2", g:"V", t:"Nền tảng theo dõi kiến nghị KTGS/KTNB của PVEP", out:"Công cụ số hóa theo dõi kiến nghị + báo cáo tình hình thực hiện hằng quý", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"KSNB · các Ban chủ trì KTGS", co:"CN-CĐS", dl:"Quý IV/2026", s:2, e:6, p:0,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"Đã review 2 dashboard hiện có ở lớp đầu việc."}], criteria:["E3"]},
    {id:"5.3", g:"V", t:"Nền tảng theo dõi kiến nghị của Petrovietnam/cấp thẩm quyền", out:"Công cụ số hóa theo dõi kiến nghị PVN + báo cáo hằng quý", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"Các Ban chủ trì lĩnh vực chuyên môn", co:"CN-CĐS", dl:"Quý IV/2026", s:2, e:6, p:0,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"Chưa bắt đầu."}], criteria:["E3"]},
    {id:"6.1", g:"VI", t:"Xây dựng cơ chế quản trị ESG của PVEP", out:"Ban hành quy định về quản trị ESG (tích hợp rủi ro khí hậu, lộ trình chuyển dịch năng lượng)", lead:"TV HĐTV Đinh Trọng Huy · PTGĐ Hoàng Ngọc Trung", own:"CL&QTRR", co:"Các Ban/VP", dl:"Quý IV/2026", s:2, e:6, p:0,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"Chưa bắt đầu."}], criteria:["C1"]},
    {id:"6.2", g:"VI", milestone:true, t:"Hoàn thiện Báo cáo ESG theo chuẩn mực (định hướng ISSB IFRS S1/S2)", out:"Báo cáo ESG 2027 chuẩn hóa, công bố cùng thời điểm với Báo cáo thường niên", lead:"TV HĐTV Đinh Trọng Huy · PTGĐ Hoàng Ngọc Trung", own:"CL&QTRR", co:"VP", dl:"Tháng 4/2027", s:4, e:10, p:16.7,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"16,7% — Báo cáo ESG 2024 theo GRI đã công bố; đang định hướng ISSB."}],
     criteria:["C2"]},
    {id:"7.1", g:"VII", t:"Cơ chế công bố thông tin, minh bạch", out:"Quy định mới về CBTT; quy trình chuẩn hóa; báo cáo thường niên song ngữ công bố tháng 4 hằng năm", lead:"TV HĐTV Đinh Trọng Huy · PTGĐ Hoàng Ngọc Trung", own:"VP", co:"CL&QTRR", dl:"Quý III/2026", s:1, e:3, p:13.3,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"13,3% — đang rà soát quy chế CBTT hiện hành."}],
     criteria:["D1","D3"]},
    {id:"7.2", g:"VII", t:"Báo cáo tài chính hợp nhất theo IFRS", out:"BCTC hợp nhất theo IFRS, minh bạch báo cáo phân khúc (IFRS 8 & IAS 24), báo cáo song ngữ", lead:"TV HĐTV Đinh Văn Đức · PTGĐ Hoàng Ngọc Trung", own:"KTKT", co:"—", dl:"Từ Quý III/2026 (liên tục)", s:1, e:10, cont:true, p:100,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"Đã vận hành BCTC hợp nhất theo IFRS — duy trì định kỳ."}],
     criteria:["D2"]},
    {id:"7.3", g:"VII", milestone:true, t:"Triển khai ERP GĐ1 & kết nối hệ thống báo cáo quản trị", out:"Hệ thống ERP Giai đoạn 1 vận hành đúng hạn", lead:"CT HĐTV Trương Quốc Lâm · PTGĐ Ngô Khánh Xạ", own:"CN-CĐS", co:"Các Ban/VP, Đơn vị", dl:"Tháng 12/2026", s:1, e:6, p:0,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"Theo kế hoạch ERP GĐ1 — mốc vận hành tháng 12/2026."}], criteria:["D2"]},
    {id:"8.1", g:"VIII", t:"Hợp nhất, hoàn thiện khung quản trị Người đại diện", out:"Ban hành Quy chế quản lý Người đại diện (ma trận năng lực, khung KPI gắn kết quả quản trị)", lead:"TV HĐTV-TGĐ Nguyễn Thiện Bảo · PTGĐ Hoàng Xuân Dương", own:"QTNNL", co:"QLDA", dl:"Quý IV/2026", s:2, e:6, p:10,
     prereq:"", issues:"", requests:"", history:[{d:"10/07/2026", note:"10% — đang tổng hợp khung quản trị NĐD hiện hành."}],
     criteria:["E1"]},
  ];

  /* ---------- CRITERIA: 14 tiêu chí cốt lõi (v1) / 5 trụ cột ACGS — mở rộng được lên 152 ----------
     status: "dat" | "dangsua" | "hut" · weight 1–3 (3 = trọng yếu) — PH5 toggle, HUD tính live.  */
  const PILLARS = {
    A:{name:"Quyền của cổ đông/chủ sở hữu", score2026:100, target2027:100},
    B:{name:"Đối xử công bằng, kiểm soát giao dịch liên quan", score2026:100, target2027:100},
    C:{name:"Bên hữu quan & phát triển bền vững (ESG)", score2026:71.2, target2027:90},
    D:{name:"Công bố thông tin & minh bạch", score2026:87.3, target2027:95},
    E:{name:"Trách nhiệm của HĐTV", score2026:88.9, target2027:95},
  };
  const CRITERIA = [
    {id:"A1", pillar:"A", name:"Thực hiện đầy đủ, đúng hạn nghĩa vụ với chủ sở hữu (cổ tức/lợi nhuận nộp)", w:2, score:5,
     action:"Duy trì; lưu bằng chứng nộp đúng hạn hằng năm", ban:"KTKT", dl:"Duy trì", tasks:[], evidence:["EV-15"]},
    {id:"A2", pillar:"A", name:"Cơ chế xin ý kiến/báo cáo chủ sở hữu (Petrovietnam) đầy đủ, kịp thời", w:2, score:5,
     action:"Duy trì kênh báo cáo NĐD; chuẩn hóa hồ sơ trình", ban:"VP · KSNB", dl:"Duy trì", tasks:[], evidence:["EV-12","EV-15"]},
    {id:"B1", pillar:"B", name:"Kiểm soát giao dịch với bên liên quan (RPT): thẩm quyền phê duyệt, giám sát", w:3, score:5,
     action:"Duy trì theo Điều lệ (Điều 36) và quy chế chuyên ngành", ban:"KSNB", dl:"Duy trì", tasks:[], evidence:["EV-15"]},
    {id:"B2", pillar:"B", name:"Phòng ngừa xung đột lợi ích, giao dịch nội gián (quy tắc đạo đức, kiêm nhiệm)", w:2, score:5,
     action:"Duy trì Sổ tay văn hóa doanh nghiệp; bổ sung quy trình giám sát tuân thủ", ban:"QTNNL", dl:"Duy trì", tasks:[], evidence:["EV-15"]},
    {id:"C1", pillar:"C", name:"Cơ chế quản trị ESG được ban hành (quy định, phân công giám sát cấp HĐTV)", w:3, score:1,
     action:"Xây dựng và ban hành quy định quản trị ESG", ban:"CL&QTRR", dl:"Quý IV/2026", tasks:["6.1"], evidence:["EV-14"]},
    {id:"C2", pillar:"C", name:"Báo cáo phát triển bền vững theo chuẩn quốc tế (GRI → ISSB IFRS S1/S2)", w:2, score:3,
     action:"Nâng Báo cáo ESG theo định hướng ISSB, kiểm định độc lập", ban:"CL&QTRR", dl:"Tháng 4/2027", tasks:["6.2"], evidence:["EV-19"]},
    {id:"C3", pillar:"C", name:"Kênh đối thoại bên hữu quan & cơ chế tố giác (whistleblowing) công bố rõ", w:1, score:3,
     action:"Hoàn thiện Whistleblowing Policy & kênh phản ánh riêng, công bố trên website", ban:"CL&QTRR · VP", dl:"Quý IV/2026", tasks:["6.1"], evidence:["EV-19"]},
    {id:"D1", pillar:"D", name:"Quy định công bố thông tin mới: danh mục, quy trình, trách nhiệm", w:3, score:3,
     action:"Ban hành quy định CBTT thay thế quy chế hiện hành", ban:"VP", dl:"Quý III/2026", tasks:["7.1","2.1"], evidence:["EV-16"]},
    {id:"D2", pillar:"D", name:"BCTC hợp nhất theo IFRS, minh bạch phân khúc & giao dịch liên quan", w:3, score:5,
     action:"Duy trì; kết nối ERP GĐ1 với hệ thống báo cáo quản trị", ban:"KTKT · CN-CĐS", dl:"Tháng 12/2026", tasks:["7.2","7.3"], evidence:["EV-15"]},
    {id:"D3", pillar:"D", name:"Báo cáo thường niên song ngữ công bố đúng hạn (tháng 4) + website đầy đủ", w:2, score:3,
     action:"Chuẩn hóa BCTN song ngữ, bổ sung mục quan hệ nhà đầu tư trên website", ban:"VP", dl:"Tháng 4/2027", tasks:["7.1"], evidence:["EV-20"]},
    {id:"E1", pillar:"E", name:"Khung quản trị HĐTV: quy chế làm việc, mô hình ủy ban, quy chế NĐD", w:3, score:3,
     action:"Nghiên cứu mô hình ủy ban; ban hành Quy chế quản lý Người đại diện", ban:"QTNNL", dl:"Quý IV/2026", tasks:["1.1","1.2","2.1","8.1"], evidence:["EV-14","EV-16"]},
    {id:"E2", pillar:"E", name:"Hệ thống theo dõi thực thi nghị quyết, chỉ đạo của HĐTV (quản trị bằng dữ liệu)", w:2, score:3,
     action:"Vận hành dashboard theo dõi NQ/chỉ đạo + báo cáo trạng thái hằng tháng", ban:"KSNB · CL&QTRR", dl:"Quý III/2026", tasks:["1.3","3.2"], evidence:["EV-14"]},
    {id:"E3", pillar:"E", name:"KTGS & KTNB theo chuẩn mực quốc tế (IIA), theo dõi kiến nghị tập trung", w:3, score:3,
     action:"Chuẩn hóa quy trình IIA; nền tảng theo dõi kiến nghị KTGS/KTNB", ban:"KSNB", dl:"Quý IV/2026", tasks:["5.1","5.2","5.3"], evidence:["EV-14"]},
    {id:"E4", pillar:"E", name:"QTRR bao quát rủi ro trọng yếu (tài chính, ESG, an ninh mạng) + KRI cảnh báo sớm", w:3, score:3,
     action:"Ban hành Quy chế QTRR sửa đổi + danh mục rủi ro + bộ KRI", ban:"CL&QTRR", dl:"Quý IV/2026", tasks:["4.1","4.2","4.3"], evidence:["EV-14"]},
  ];

  /* ---------- EVIDENCE: văn bản & bằng chứng (label: "Tham chiếu" | "Bằng chứng") ---------- */
  const EVIDENCE = [
    {id:"EV-01", name:"Nghị quyết 79-NQ/TW — Bộ Chính trị (06/01/2026)", type:"Chỉ đạo TW", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/TW/Nghị quyết 79-NQ.TW về phát triển kinh tế NN.pdf"),
     note:"Mục tiêu 2030: 100% TĐ/TCT nhà nước áp dụng quản trị OECD (Mục II.2).", criteria:[], tasks:[]},
    {id:"EV-02", name:"Nghị quyết 01/NQ-CNNL — HĐTV Petrovietnam (10/01/2026)", type:"Chỉ đạo PVN", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVN/20260110_01.NQ-CNNL_ve nhiem vu giai phap trong tam thuc hien Ke hoach nam 2026.pdf"),
     note:"OECD là mục tiêu trung — dài hạn của Tập đoàn.", criteria:[], tasks:[]},
    {id:"EV-03", name:"Kế hoạch 12-KH/ĐU — Đảng ủy Petrovietnam (30/01/2026)", type:"Chỉ đạo PVN", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVN/20260130_12-KH.DU thực hiện CT2 - Hoàn thiện thể chế. cơ chế tạo điều kiện cho DN chủ động và phát triển.pdf"),
     note:"Tiếp cận chọn lọc mô hình quản trị tiên tiến của OECD.", criteria:[], tasks:[]},
    {id:"EV-04", name:"CV 2971/CNNL-HĐTV — HĐTV Petrovietnam (16/04/2026)", type:"Chỉ đạo PVN", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVN/20260416_2971.CNNL-HĐTV_Trien khai ap dung cac nguyen tac quan tri cong ty cua OECD tai doanh nghiep.pdf"),
     note:"Yêu cầu NĐD PVEP rà soát, đánh giá thực trạng quản trị theo OECD.", criteria:["A2"], tasks:[]},
    {id:"EV-05", name:"CV 3698/CNNL-KSNB — Ban KSNB Petrovietnam (11/05/2026)", type:"Chỉ đạo PVN", label:"Tham chiếu",
     link:spO("5. Bo cau hoi khao sat cham diem/260515 Trình HĐTV/20260511_3698.CNNL-KSNB_Tu ra soat danh gia thuc trang quan tri DN theo OECD.pdf"),
     note:"Yêu cầu tự chấm điểm theo Bộ câu hỏi khảo sát, gửi trước 15/05/2026.", criteria:["A2"], tasks:[]},
    {id:"EV-06", name:"Nghị quyết 5432/NQ-CNNL — HĐTV Petrovietnam (30/06/2026)", type:"Chỉ đạo PVN", label:"Tham chiếu",
     link:spO("8. KH trien khai ap dung OECD tai PVEP/20260630_5432.NQ-CNNL_KH trien khai ap dung nguyen tac quan tri cong ty OECD.pdf"),
     note:"Kế hoạch triển khai OECD tại Petrovietnam & ĐVTV; yêu cầu báo cáo định kỳ (mục 2.8).", criteria:[], tasks:[]},
    {id:"EV-07", name:"Kế hoạch 10-KH/ĐU — Đảng ủy PVEP (11/02/2026)", type:"Chỉ đạo PVEP", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/260211_10-KH-DU Trien khai thuc hien CTr-03 va KH trien khai cac chuong trinh thanh phan thuc hien NQ Dang bo PVN 2025-2030.pdf"),
     note:"Hoàn thành Bộ quy chế quản trị nội bộ áp dụng nguyên tắc OECD.", criteria:["E1"], tasks:["2.1"]},
    {id:"EV-08", name:"CTHĐ 09-CTr/ĐU — Đảng ủy PVEP (25/02/2026)", type:"Chỉ đạo PVEP", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/20260225_09.CTr-ĐU_Chuong trinh hanh dong NQ Dai hoi Dang bo TCT lan thu IV 2025-2030.pdf"),
     note:"Tối ưu hóa hệ thống quản trị hướng chuẩn OECD.", criteria:[], tasks:[]},
    {id:"EV-09", name:"CTHĐ 119/TDKT-HĐTV — HĐTV PVEP (10/02/2026)", type:"Chỉ đạo PVEP", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/20260210_119.TDKT-HĐTV_Chuong trinh hanh dong Thuc hien cac nhiem vu trong tam nam 2026 cua Tong cong ty Tham do Khai thac Dau khi.pdf"),
     note:"Hoàn thiện mô hình quản trị tiệm cận OECD 2026–2027.", criteria:[], tasks:[]},
    {id:"EV-10", name:"TBKL 130/TB-TDKT — HĐTV PVEP (12/02/2026)", type:"Chỉ đạo PVEP", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/20260212_130.TB-TDKT_Thong bao ket luan cua HDTV tai cuoc hop giao ban Hoi dong Thanh vien va Ban Tong giam doc ky hop thang 02.2026.pdf"),
     note:"Quản trị dựa trên dữ liệu, hướng tới tiêu chuẩn OECD.", criteria:["E2"], tasks:["1.3"]},
    {id:"EV-11", name:"TBKL 257/TB-TDKT — HĐTV PVEP (01/04/2026)", type:"Chỉ đạo PVEP", label:"Tham chiếu",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/20260401_257.TB-TDKT_Cuoc hop ve cong tac quan tri va quan ly doanh nghiep tai PVEP.pdf"),
     note:"Giao CL&QTRR phối hợp PVN xây lộ trình OECD trong tháng 4/2026.", criteria:[], tasks:[]},
    {id:"EV-12", name:"CV 1076/TDKT-HĐTV — HĐTV PVEP (12/05/2026)", type:"Báo cáo PVEP → PVN", label:"Bằng chứng",
     link:spO("5. Bo cau hoi khao sat cham diem/260515 Trình HĐTV/20260512_1076.TDKT-HĐTV_Trien khai ap dung cac nguyen tac quan tri doanh nghiep OECD tai PVEP.pdf"),
     note:"Báo cáo Petrovietnam kết quả rà soát theo CV 2971.", criteria:["A2"], tasks:[]},
    {id:"EV-13", name:"Biên bản 566/BB-TDKT — HĐTV PVEP (08/07/2026)", type:"Hồ sơ họp", label:"Bằng chứng",
     link:spO("8. KH trien khai ap dung OECD tai PVEP/20260708_566.BB-TDKT_Hop HDTV 07.7.2026_trien khai ap dung nguyen tac quan tri OECD.pdf"),
     note:"HĐTV thống nhất Kế hoạch triển khai OECD.", criteria:[], tasks:[]},
    {id:"EV-14", name:"QĐ 1017/QĐ-TDKT — HĐTV PVEP (10/07/2026)", type:"Quyết định", label:"Bằng chứng",
     link:spO("8. KH trien khai ap dung OECD tai PVEP/260710_QD 1017_HDTV_KH ttrien khai OECD tai PVEP.pdf"),
     note:"Ban hành KH 21 nhiệm vụ / 8 nhóm; TGĐ báo cáo HĐTV hằng tháng (mục II.3).",
     criteria:["C1","E1","E2","E3","E4"], tasks:["1.1","1.2","1.3","2.1","3.1","4.1","5.1","6.1","7.1","8.1"]},
    {id:"EV-15", name:"CV 1152/TDKT-CL&QTRR + Bộ ACGS 152 tiêu chí (20/05/2026)", type:"Kết quả tự đánh giá", label:"Bằng chứng",
     link:spO("5. Bo cau hoi khao sat cham diem/260520 Trinh PVN/260520_CV 1152_Bao cao ve ket qua ra soat OECD tai PVEP.pdf"),
     note:"Kết quả 407/456 = 89,25% — Hạng A; toàn bộ 152 câu + bằng chứng trong phụ lục.",
     criteria:["A1","A2","B1","B2","D2"], tasks:[]},
    {id:"EV-16", name:"CV 140/CVNB-HNT + Danh mục 79 VBQLNB (30/04/2026)", type:"Báo cáo cấp Ban", label:"Bằng chứng",
     link:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/BC cap Ban/20260430_140.CVNB-HNT_Hệ thống VBQLNB PVEP theo quy định pháp luật và nguyên tắc quản trị doanh nghiệp OECD.pdf"),
     link2:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/BC cap Ban/20260430_Dinh kem 140.CVNB-HNT_Danh muc he thong 79 VBQLNB PVEP cap nhat den 29.4.2026.pdf"),
     note:"Rà soát hệ thống VBQLNB theo pháp luật & OECD — đầu vào nhiệm vụ 2.1.", criteria:["D1","E1"], tasks:["2.1"]},
    {id:"EV-17", name:"CV 169/CVNB-HNT — Báo cáo rà soát thực trạng (18/05/2026)", type:"Báo cáo cấp Ban", label:"Bằng chứng",
     link:spO("5. Bo cau hoi khao sat cham diem/260515 Trình HĐTV/20260518_169.CVNB-HNT_Bao cao Ra soat danh gia thuc trang quan tri doanh nghiep theo OECD.pdf"),
     note:"Báo cáo kết quả rà soát trình Ban TGĐ.", criteria:[], tasks:[]},
    {id:"EV-18", name:"Tờ trình 99/NB-KSNB — trình ban hành KH OECD (08/07/2026)", type:"Tờ trình", label:"Bằng chứng",
     link:spO("4. Memo Ban KSNB/20260708_99.NB-KSNB_KH trien khai ap dung cac nguyen tac quan tri OECD.pdf"),
     note:"Tờ trình của Trưởng Ban KSNB — cơ sở ban hành QĐ 1017.", criteria:[], tasks:[]},
    {id:"EV-19", name:"Báo cáo ESG 2024 của PVEP (chuẩn GRI, kiểm định E&Y)", type:"Báo cáo công bố", label:"Bằng chứng",
     link:"https://www.pvep.com.vn", note:"Công bố trên website PVEP — bằng chứng trụ cột ESG (ACGS B.1).",
     criteria:["C2","C3"], tasks:["6.2"]},
    {id:"EV-20", name:"Báo cáo thường niên 2025 & website công bố thông tin PVEP", type:"Báo cáo công bố", label:"Bằng chứng",
     link:"https://www.pvep.com.vn", note:"BCTN công bố 04/05/2026 trên 1-Office & website — bằng chứng CBTT (ACGS C.8).",
     criteria:["D3"], tasks:["7.1"]},
  ];

  /* ---------- DIRECTIVES: ma trận lịch chỉ đạo PH4 (level × tháng 2026) ----------
     level: TW | PVN | DU | HDTV · month: tháng ban hành · parentId: văn bản cấp trên
     reportedBy: số hiệu văn bản PVEP đã báo cáo/cụ thể hóa · milestone: mốc then chốt */
  const LEVELS = {
    TW:  {name:"Bộ Chính trị / TW",  color:"#DC2626"},
    PVN: {name:"Petrovietnam",       color:"#23286b"},
    DU:  {name:"Đảng ủy PVEP",       color:"#D97706"},
    HDTV:{name:"HĐTV PVEP",          color:"#046B39"},
  };
  const DIRECTIVES = [
    {id:"nq79", level:"TW", date:"06/01/2026", month:1, ref:"NQ 79-NQ/TW", org:"Bộ Chính trị", milestone:true, parentId:null, reportedBy:null,
     summary:"Mục tiêu 2030: 100% tập đoàn, TCT nhà nước áp dụng nguyên tắc quản trị OECD",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/TW/Nghị quyết 79-NQ.TW về phát triển kinh tế NN.pdf"),
     kp:["Nghị quyết của Bộ Chính trị về phát triển kinh tế nhà nước — văn bản gốc khởi nguồn chương trình OECD.","Mục tiêu cụ thể đến 2030: 100% tập đoàn kinh tế, tổng công ty nhà nước áp dụng nguyên tắc quản trị của OECD (Mục II.2, trang 4)."]},
    {id:"nq01", level:"PVN", date:"10/01/2026", month:1, ref:"NQ 01/NQ-CNNL", org:"HĐTV Petrovietnam", parentId:"nq79", reportedBy:null,
     summary:"OECD là mục tiêu trung — dài hạn; đồng bộ quản trị hiện đại: chuỗi giá trị, số, rủi ro, tuân thủ",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/PVN/20260110_01.NQ-CNNL_ve nhiem vu giai phap trong tam thuc hien Ke hoach nam 2026.pdf"),
     kp:["Xác định áp dụng chuẩn OECD là mục tiêu trung hạn và dài hạn của Tập đoàn (Mục II.2.iv).","Chuẩn hóa quy trình, phân cấp — phân quyền gắn nâng cao hiệu quả kiểm soát."]},
    {id:"kh12", level:"PVN", date:"30/01/2026", month:1, ref:"KH 12-KH/ĐU", org:"Đảng ủy Petrovietnam", parentId:"nq79", reportedBy:null,
     summary:"Chương trình 2 — hoàn thiện thể chế; tiếp cận chọn lọc mô hình quản trị tiên tiến OECD",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/PVN/20260130_12-KH.DU thực hiện CT2 - Hoàn thiện thể chế. cơ chế tạo điều kiện cho DN chủ động và phát triển.pdf"),
     kp:["Thể chế hóa kịp thời nghị quyết của Trung ương, Bộ Chính trị.","Tiếp cận có chọn lọc thông lệ, mô hình quản trị tiên tiến của OECD."]},
    {id:"cthd119", level:"HDTV", date:"10/02/2026", month:2, ref:"CTHĐ 119/TDKT-HĐTV", org:"HĐTV PVEP", parentId:"nq01", reportedBy:null,
     summary:"Hoàn thiện mô hình quản trị doanh nghiệp tiệm cận chuẩn OECD trong 2026–2027",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/20260210_119.TDKT-HĐTV_Chuong trinh hanh dong Thuc hien cac nhiem vu trong tam nam 2026 cua Tong cong ty Tham do Khai thac Dau khi.pdf"),
     kp:["Quan điểm trọng tâm 2026: tối ưu hóa hệ thống quản trị hướng chuẩn OECD (Mục II.2.iv).","Phân công CT HĐTV & TGĐ chỉ đạo; QTNNL, CL&QTRR thực hiện trong năm 2026."]},
    {id:"kh10", level:"DU", date:"11/02/2026", month:2, ref:"KH 10-KH/ĐU", org:"Đảng ủy PVEP", parentId:"kh12", reportedBy:null,
     summary:"Hoàn thành Bộ quy chế quản trị nội bộ áp dụng nguyên tắc OECD; phân quyền kèm KTGS",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/260211_10-KH-DU Trien khai thuc hien CTr-03 va KH trien khai cac chuong trinh thanh phan thuc hien NQ Dang bo PVN 2025-2030.pdf"),
     kp:["Hiện đại hóa quản trị, hoàn thành Bộ quy chế quản trị nội bộ theo OECD (Mục II.2.2).","Tăng phân quyền chủ động kèm cơ chế kiểm tra giám sát, QTRR theo OECD."]},
    {id:"tb130", level:"HDTV", date:"12/02/2026", month:2, ref:"TBKL 130/TB-TDKT", org:"HĐTV PVEP", parentId:null, reportedBy:null,
     summary:"Quản trị dựa trên dữ liệu, minh bạch, tăng cường QTRR — hướng tới tiêu chuẩn OECD",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/20260212_130.TB-TDKT_Thong bao ket luan cua HDTV tai cuoc hop giao ban Hoi dong Thanh vien va Ban Tong giam doc ky hop thang 02.2026.pdf"),
     kp:["Kết luận giao ban HĐTV–BTGĐ tháng 02/2026 (Mục 1).","Nâng cao năng lực quản trị dựa trên dữ liệu, bổ sung hệ thống đánh giá hiệu quả."]},
    {id:"ctr09", level:"DU", date:"25/02/2026", month:2, ref:"CTHĐ 09-CTr/ĐU", org:"Đảng ủy PVEP", parentId:"kh12", reportedBy:null,
     summary:"Tối ưu hóa hệ thống quản trị, từng bước hướng tới quản trị theo tiêu chuẩn OECD",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/20260225_09.CTr-ĐU_Chuong trinh hanh dong NQ Dai hoi Dang bo TCT lan thu IV 2025-2030.pdf"),
     kp:["Chương trình hành động thực hiện NQ Đại hội Đảng bộ PVEP lần IV, 2025–2030.","Định hướng tối ưu hóa hệ thống quản trị theo tiêu chuẩn OECD (Mục II.8 — PL 2.11)."]},
    {id:"tb257", level:"HDTV", date:"01/04/2026", month:4, ref:"TBKL 257/TB-TDKT", org:"HĐTV PVEP", parentId:"nq01", reportedBy:null,
     summary:"Giao CL&QTRR phối hợp Petrovietnam xây lộ trình áp dụng mô hình OECD trong tháng 4/2026",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/PVEP/20260401_257.TB-TDKT_Cuoc hop ve cong tac quan tri va quan ly doanh nghiep tai PVEP.pdf"),
     kp:["Kết luận họp chuyên đề về công tác quản trị, quản lý doanh nghiệp.","Lộ trình phù hợp định hướng Tập đoàn, đáp ứng nguyên tắc OECD."]},
    {id:"cv2971", level:"PVN", date:"16/04/2026", month:4, ref:"CV 2971/CNNL-HĐTV", org:"HĐTV Petrovietnam", parentId:"nq79",
     reportedBy:"CV 1076/TDKT-HĐTV (12/05/2026)",
     summary:"NĐD PVEP tổ chức rà soát, đánh giá thực trạng quản trị doanh nghiệp theo nguyên tắc OECD",
     fileUrl:spO("3. Tap hop VB chi dao OECD cac cap/PVN/20260416_2971.CNNL-HĐTV_Trien khai ap dung cac nguyen tac quan tri cong ty cua OECD tai doanh nghiep.pdf"),
     kp:["Chỉ đạo trực tiếp tới Người đại diện vốn tại PVEP về triển khai OECD.","Căn cứ để PVEP thực hiện bộ khảo sát ACGS 152 câu."]},
    {id:"cv3698", level:"PVN", date:"11/05/2026", month:5, ref:"CV 3698/CNNL-KSNB", org:"Ban KSNB Petrovietnam", parentId:"cv2971",
     reportedBy:"CV 1152/TDKT-CL&QTRR (20/05/2026)",
     summary:"Tự chấm điểm theo Bộ câu hỏi khảo sát; gửi kết quả về Petrovietnam trước 15/05/2026",
     fileUrl:spO("5. Bo cau hoi khao sat cham diem/260515 Trình HĐTV/20260511_3698.CNNL-KSNB_Tu ra soat danh gia thuc trang quan tri DN theo OECD.pdf"),
     kp:["TL. Chủ tịch HĐTV PVN (Trưởng Ban KSNB Bùi Thị Nguyệt ký) — tiếp nối CV 2971.","PVEP hoàn thành đúng hạn: gửi 14/05, kết quả 89,25% Hạng A."]},
    {id:"cv1076", level:"HDTV", date:"12/05/2026", month:5, ref:"CV 1076/TDKT-HĐTV", org:"HĐTV PVEP", parentId:"cv2971", reportedBy:null,
     summary:"Báo cáo Petrovietnam kết quả tự rà soát, đánh giá thực trạng quản trị theo bộ tiêu chí OECD",
     fileUrl:spO("5. Bo cau hoi khao sat cham diem/260515 Trình HĐTV/20260512_1076.TDKT-HĐTV_Trien khai ap dung cac nguyen tac quan tri doanh nghiep OECD tai PVEP.pdf"),
     kp:["Văn bản báo cáo Petrovietnam theo yêu cầu CV 2971.","Kèm kết quả ACGS 407/456 = 89,25% — Hạng A."]},
    {id:"nq5432", level:"PVN", date:"30/06/2026", month:6, ref:"NQ 5432/NQ-CNNL", org:"HĐTV Petrovietnam", milestone:true, parentId:"cv2971",
     reportedBy:"QĐ 1017/QĐ-TDKT (10/07/2026)",
     summary:"Kế hoạch triển khai áp dụng nguyên tắc quản trị công ty OECD tại Petrovietnam & ĐVTV",
     fileUrl:spO("8. KH trien khai ap dung OECD tai PVEP/20260630_5432.NQ-CNNL_KH trien khai ap dung nguyen tac quan tri cong ty OECD.pdf"),
     kp:["Kế hoạch toàn Tập đoàn; NĐD kiểm soát định kỳ báo cáo PVN (mục 2.8).","Đánh giá chéo định kỳ; đánh giá độc lập quý II/2027 (mục 2.3)."]},
    {id:"bb566", level:"HDTV", date:"08/07/2026", month:7, ref:"BB 566/BB-TDKT", org:"HĐTV PVEP", parentId:"nq5432", reportedBy:null,
     summary:"Họp HĐTV thống nhất Kế hoạch triển khai OECD tại PVEP theo NQ 5432",
     fileUrl:spO("8. KH trien khai ap dung OECD tai PVEP/20260708_566.BB-TDKT_Hop HDTV 07.7.2026_trien khai ap dung nguyen tac quan tri OECD.pdf"),
     kp:["Thống nhất trên cơ sở Tờ trình 99/NB-KSNB của Ban KSNB.","Hai ngày sau HĐTV ban hành QĐ 1017."]},
    {id:"qd1017", level:"HDTV", date:"10/07/2026", month:7, ref:"QĐ 1017/QĐ-TDKT", org:"HĐTV PVEP", milestone:true, parentId:"nq5432", reportedBy:null,
     summary:"Ban hành Kế hoạch triển khai OECD tại PVEP — 21 nhiệm vụ / 8 nhóm, phân công đến 2027",
     fileUrl:spO("8. KH trien khai ap dung OECD tai PVEP/260710_QD 1017_HDTV_KH ttrien khai OECD tai PVEP.pdf"),
     kp:["Mục tiêu: quản trị theo chuẩn OECD 2026, hàng đầu quốc gia 2026–2030.","TGĐ báo cáo HĐTV kết quả thực hiện định kỳ hằng tháng (Mục II.3)."]},
  ];

  /* ---------- SCORES theo năm (thực tế + mục tiêu — KHÔNG bịa số quá khứ) ---------- */
  const SCORES = [
    {y:"2024", v:null,  note:"Chưa thực hiện đánh giá ACGS"},
    {y:"2025", v:null,  note:"Chưa thực hiện đánh giá ACGS"},
    {y:"2026", v:89.25, note:"Tự đánh giá 152 tiêu chí — CV 1152 (20/05/2026)", real:true},
    {y:"2027", v:95,    note:"Mục tiêu: quản trị hàng đầu quốc gia (QĐ 1017 · Mục I.2)", target:true},
  ];

  /* ---------- DOCS PH1 (giữ nguyên dữ liệu & link từ index gốc) ---------- */
  const DOCS = {
    original: [
      {t:"G20/OECD Principles of Corporate Governance", org:"OECD · G20", yr:"2023", d:"Bộ nguyên tắc gốc — 6 trụ cột quản trị công ty, nền tảng của toàn bộ chương trình.", f:"PDF", u:spO("1. Tài liệu từ tổ chức OECD/G20.OECD Principles of Corporate Governance.pdf"),
       kp:["Chuẩn mực quản trị công ty toàn cầu do G20 và OECD thông qua tháng 9/2023 (bản cập nhật từ phiên bản 2015).","6 chương: (I) Nền tảng khung quản trị hiệu quả · (II) Quyền và đối xử công bằng với cổ đông · (III) Nhà đầu tư tổ chức, thị trường và trung gian · (IV) Công bố thông tin & minh bạch · (V) Trách nhiệm của Hội đồng quản trị · (VI) Phát triển bền vững & sức chống chịu.","Điểm mới 2023: chương riêng về bền vững — khí hậu, ESG; nhấn mạnh số hóa, an ninh mạng và vai trò HĐQT trong giám sát rủi ro.","Là cột \"Tham chiếu OECD 2023\" trong bộ chấm điểm ACGS 152 câu của PVEP."]},
      {t:"OECD Guidelines on Corporate Governance of SOEs", org:"OECD", yr:"2024", d:"Hướng dẫn quản trị dành riêng cho doanh nghiệp nhà nước — khung tham chiếu chính của PVEP.", f:"PDF", u:spO("1. Tài liệu từ tổ chức OECD/OECD Guidelines on Corporate Governance of SOEs.pdf"),
       kp:["Khuyến nghị của OECD dành riêng cho DNNN, bản cập nhật 2024 (từ phiên bản 2015).","7 chương: lý do sở hữu nhà nước · nhà nước với vai trò chủ sở hữu · DNNN trên thị trường (sân chơi bình đẳng) · đối xử công bằng với cổ đông/nhà đầu tư · quan hệ bên hữu quan & kinh doanh có trách nhiệm · công bố thông tin & minh bạch · trách nhiệm HĐQT DNNN.","Bản 2024 bổ sung nội dung khí hậu, chuyển dịch năng lượng, liêm chính và chống tham nhũng.","Khung tham chiếu sát nhất với PVEP — doanh nghiệp 100% vốn thuộc Petrovietnam."]},
      {t:"Review of the Corporate Governance of SOEs in Vietnam", org:"OECD", yr:"2022", d:"Báo cáo OECD rà soát quản trị DNNN Việt Nam — bối cảnh và khuyến nghị cho Việt Nam.", f:"PDF", u:spO("1. Tài liệu từ tổ chức OECD/OECD Review of the corporate governance of SOEs in Vietnam.pdf"),
       kp:["OECD rà soát khung quản trị DNNN Việt Nam theo bộ SOE Guidelines.","Nhận diện các điểm yếu: chức năng chủ sở hữu nhà nước còn phân tán, tính độc lập và chuyên nghiệp của HĐQT/HĐTV hạn chế, công bố thông tin chưa theo chuẩn quốc tế.","Khuyến nghị: tập trung hóa chức năng chủ sở hữu, chuyên nghiệp hóa bộ máy quản trị, nâng chuẩn minh bạch báo cáo tài chính.","Là bối cảnh trực tiếp dẫn tới chủ trương 100% tập đoàn, tổng công ty nhà nước áp dụng OECD tại NQ 79-NQ/TW."]},
      {t:"Báo cáo đánh giá quản trị công ty DNNN Việt Nam", org:"OECD · bản tiếng Việt", yr:"VN", vn:true, d:"Bản tiếng Việt của báo cáo đánh giá — thuận tiện tra cứu nhanh cho Lãnh đạo.", f:"PDF", u:spO("1. Tài liệu từ tổ chức OECD/Báo cáo đánh giá quản trị công ty DNNN Viet Nam.pdf"),
       kp:["Bản dịch tiếng Việt của báo cáo OECD đánh giá quản trị DNNN Việt Nam.","Phục vụ Lãnh đạo đọc nhanh và thống nhất thuật ngữ tiếng Việt khi soạn văn bản nội bộ.","Dùng kèm bản gốc tiếng Anh khi cần trích dẫn chính xác."]},
      {t:"OECD Corporate Governance Factbook 2025", org:"OECD", yr:"2025", d:"Dữ liệu so sánh khung quản trị của 49 nền kinh tế — phiên bản mới nhất.", f:"PDF", u:spO("1. Tài liệu từ tổ chức OECD/OECD Corporate Governance Factbook 2025.pdf"),
       kp:["Dữ liệu so sánh khung quản trị công ty của khoảng 49 nền kinh tế, cập nhật 2 năm/lần.","Bao quát: cấu trúc sở hữu, khung pháp lý, ĐHĐCĐ và quyền cổ đông, cơ cấu — thù lao HĐQT, kiểm toán, công bố ESG.","Dùng làm benchmark khi PVEP xây dựng, sửa đổi quy chế quản trị nội bộ."]},
      {t:"OECD Corporate Governance Factbook 2023", org:"OECD", yr:"2023", d:"Phiên bản 2023 — dùng đối chiếu xu hướng thay đổi giữa hai kỳ.", f:"PDF", u:spO("1. Tài liệu từ tổ chức OECD/OECD Corporate Governance Factbook 2023.pdf"),
       kp:["Phiên bản kỳ trước của Factbook (2023).","Dùng đối chiếu với bản 2025 để thấy xu hướng thay đổi chính sách quản trị giữa hai kỳ."]},
      {t:"Ownership and Governance of State-Owned Enterprises", org:"OECD", yr:"2024", d:"Tổng hợp thông lệ sở hữu & quản trị DNNN các nước thành viên và đối tác OECD.", f:"PDF", u:spO("1. Tài liệu từ tổ chức OECD/Ownership and Governance of SOEs 2024.pdf"),
       kp:["Tuyển tập thông lệ quốc gia (compendium) về sở hữu và quản trị DNNN của các nước OECD và đối tác.","So sánh các mô hình cơ quan chủ sở hữu: tập trung, hỗn hợp, phân tán; quy trình đề cử — bổ nhiệm HĐQT; cơ chế giám sát hiệu quả.","Tham khảo trực tiếp khi thiết kế quan hệ Petrovietnam – PVEP và cơ chế Người đại diện vốn (nhóm nhiệm vụ VIII của KH 1017)."]},
    ],
    training: [
      {t:"Bộ Nguyên tắc quản trị công ty Việt Nam 2026 (VN CG Code)", org:"UBCKNN · IFC", yr:"2026", vn:true, d:"Chuẩn quản trị công ty Việt Nam cập nhật 2026 — cầu nối giữa quy định trong nước và OECD.", f:"PDF", u:spO("2. Tài liệu tham khảo/260423_Bo Nguyen tac quan tri cong ty VN 2026 (VN CG CODE 2026).pdf"),
       kp:["Bộ nguyên tắc quản trị công ty của Việt Nam, bản cập nhật 2026, tiệm cận Nguyên tắc OECD 2023.","Áp dụng theo cơ chế \"tuân thủ hoặc giải thích\" — cao hơn mức tối thiểu của pháp luật.","Là cầu nối giữa quy định trong nước (Luật Doanh nghiệp, NĐ 155…) và chuẩn mực OECD khi PVEP rà soát VBQLNB."]},
      {t:"Tài liệu Hội thảo Quản trị công ty của Petrovietnam", org:"Petrovietnam", yr:"01/2026", vn:true, d:"Tài liệu hội thảo QTCT toàn Tập đoàn (cập nhật 20/01/2026) — định hướng chung cho các ĐVTV.", f:"PDF", u:spO("2. Tài liệu tham khảo/PVN_Tài liệu Hội thảo về QTCT_updated 20260120.pdf"),
       kp:["Tài liệu hội thảo quản trị công ty toàn Tập đoàn (cập nhật 20/01/2026).","Truyền đạt định hướng của Petrovietnam về áp dụng OECD cho các đơn vị thành viên.","Tài liệu đào tạo nền cho đội ngũ đầu mối OECD của các Ban/Đơn vị (nhiệm vụ 1.1 KH 1017)."]},
      {t:"CGS/VNICG — Đánh giá OECD tại Petrovietnam", org:"CGS · Hội thảo PVN", yr:"05/2026", vn:true, d:"Bài giới thiệu của công ty tư vấn CGS về phương pháp đánh giá theo OECD (hội thảo 20/05/2026).", f:"PDF", u:spO("2. Tài liệu tham khảo/6-Gioi thieu cua cong ty CGS ve OECD_Hoi thao PVN_201526/CGS_VNICG_PVN_OECD assessment_20260520.pdf"),
       kp:["Bài trình bày của tư vấn CGS/VNICG tại hội thảo Petrovietnam ngày 20/05/2026.","Giới thiệu phương pháp đánh giá mức độ đáp ứng OECD theo thang ACGS (ASEAN Corporate Governance Scorecard).","Cơ sở phương pháp luận cho bộ chấm điểm Hybrid 152 câu mà Ban KSNB xây dựng cho PVEP."]},
      {t:"Bộ câu hỏi & chấm điểm ACGS Hybrid — 152 tiêu chí", org:"Ban KSNB · PVEP", yr:"05/2026", vn:true, d:"Bộ công cụ tự đánh giá theo ASEAN CG Scorecard — kết quả PVEP đạt 89,25%, Hạng A.", f:"PDF", u:spO("5. Bo cau hoi khao sat cham diem/260520 Trinh PVN/260520_CV 1152_Bao cao ve ket qua ra soat OECD tai PVEP.pdf"),
       kp:["152 câu hỏi theo ACGS, chấm Hybrid: nhóm Có/Không (0 hoặc 3 điểm) và nhóm thang 0–3 theo 4 rubric; Level 2 có 18 câu thưởng (+1) và 26 câu phạt (−1/−2).","Kết quả PVEP: 407/456 = 89,25% — Hạng A (A: 100% · B: 71,2% · C: 87,3% · D: 88,9%); 4 điểm thưởng, không điểm phạt.","Hoàn thiện qua 4 vòng: phối hợp CL&QTRR (14/05) → trình HĐTV (15/05) → tiếp thu ý kiến (19/05) → gửi Petrovietnam kèm CV 1152 (20/05).","Bản mở kèm theo CV 1152/TDKT-CL&QTRR — toàn bộ 152 câu, điểm và bằng chứng nằm trong phụ lục công văn."]},
      {t:"King IV Report on Corporate Governance", org:"IoDSA · Nam Phi", yr:"2016", d:"Bộ quy tắc quản trị nổi tiếng theo nguyên tắc 'apply and explain' — tham khảo quốc tế ngoài OECD.", f:"PDF", u:"",
       kp:["Bộ quy tắc quản trị công ty của Nam Phi (Institute of Directors South Africa), áp dụng 'apply and explain' cho mọi loại hình tổ chức kể cả DNNN.","17 nguyên tắc xoay quanh 4 kết quả quản trị: văn hóa đạo đức, hiệu quả hoạt động, kiểm soát hữu hiệu và tính chính danh.","Nguồn tham khảo tốt cho quản trị ESG, vai trò ủy ban và báo cáo tích hợp (integrated reporting).","★ Chưa có file trong kho — Nam tải bản PDF và điền link vào trường u của thẻ này."]},
      {t:"Bộ memo nghiên cứu OECD của Ban KSNB (Level 2–3 & Gap analysis)", org:"Ban KSNB", yr:"01–04/2026", vn:true, d:"Chuỗi memo nội bộ: hệ thống hóa nguyên tắc OECD, đối chiếu hiện trạng PVEP, cập nhật chỉ đạo các cấp.", f:"Nội bộ", u:"",
       kp:["Memo Level 2 (26/01) hệ thống hóa nguyên tắc OECD với nhà nước — chủ sở hữu và doanh nghiệp; Level 3 (02/02) chi tiết hóa mức thực hành.","Gap analysis (10/02) đối chiếu hiện trạng PVEP từng nguyên tắc — nhận diện khoảng trống: mô hình ủy ban HĐTV, CBTT, ESG, Người đại diện.","Báo cáo sơ bộ gửi Chủ tịch HĐTV (25/03) với 4 nội dung: chỉ đạo các cấp · nội dung OECD · đối chiếu hiện trạng · đánh giá khả năng đáp ứng.","Tài liệu làm việc nội bộ — liên hệ Ban KSNB khi cần bản đầy đủ."]},
      {t:"Báo cáo OECD tại Hội nghị Người đại diện PVN 2026", org:"Ban KSNB · PVEP", yr:"06/2026", vn:true, d:"Slide và tham luận của PVEP về kết quả triển khai OECD tại Hội nghị Người đại diện vốn Petrovietnam.", f:"PPTX", u:spO("7. Hội nghị NĐD 2026/2. Ban KSNB_BC OECD tại HN NĐD 2026_v4.pptx"),
       kp:["Bộ slide Ban KSNB báo cáo kết quả triển khai OECD tại PVEP, trình bày tại Hội nghị Người đại diện vốn Petrovietnam 2026.","Kèm tham luận về kinh nghiệm tự đánh giá theo ACGS và xây dựng kế hoạch triển khai.","PVEP thuộc nhóm đơn vị đầu tiên của Tập đoàn hoàn thành tự đánh giá đạt Hạng A.","Bản trình bày trên SharePoint M365-KSNB (yêu cầu tài khoản PVEP) — bộ slide Ban KSNB dùng tại Hội nghị Người đại diện Petrovietnam 2026."]},
      {t:"Tờ trình 99/NB-KSNB — trình HĐTV ban hành KH triển khai OECD", org:"Ban KSNB", yr:"07/2026", vn:true, d:"Tờ trình của Trưởng Ban KSNB đề xuất HĐTV phê duyệt Kế hoạch triển khai OECD tại PVEP.", f:"PDF", u:spO("4. Memo Ban KSNB/20260708_99.NB-KSNB_KH trien khai ap dung cac nguyen tac quan tri OECD.pdf"),
       kp:["Tờ trình ngày 08/07/2026, xây dựng trên cơ sở NQ 5432/NQ-CNNL và kết luận họp HĐTV (BB 566/BB-TDKT).","Đề xuất kế hoạch 21 nhiệm vụ / 8 nhóm kèm phân công lãnh đạo, Ban chủ trì – phối hợp và thời hạn.","Được HĐTV chấp thuận, ban hành thành QĐ 1017/QĐ-TDKT ngày 10/07/2026."]},
    ],
    /* VBQLNB liên quan trực tiếp OECD — khởi tạo từ sản phẩm đầu ra KH 1017.
       ★ Cập nhật so/u khi từng văn bản được ban hành. */
    vbqlnb: [
      {so:"", ten:"Bộ quy chế quản trị nội bộ áp dụng nguyên tắc OECD", loai:"Quản trị & điều hành", ban:"CL&QTRR", trangthai:"chua", u:"", task:"2.1",
       kp:["Sản phẩm tổng thể theo yêu cầu của Đảng ủy PVEP tại KH 10-KH/ĐU (hoàn thành Bộ quy chế quản trị nội bộ theo OECD).","Tổng hợp từ kết quả rà soát VBQLNB — nhiệm vụ 2.1 KH 1017 (hạn tháng 8/2026)."]},
      {so:"", ten:"Quy chế Quản trị rủi ro (sửa đổi, bổ sung theo OECD)", loai:"Kiểm soát – rủi ro", ban:"CL&QTRR", trangthai:"dangsua", u:"", task:"4.1",
       kp:["Bổ sung rủi ro tài chính, vận hành, ESG, chuỗi cung ứng, an ninh mạng, địa chính trị và đầu tư nước ngoài.","Kèm danh mục rủi ro trọng yếu cấp PVEP + bộ KRI, ngưỡng cảnh báo sớm — nhiệm vụ 4.1 KH 1017 (hạn Quý IV/2026)."]},
      {so:"", ten:"Quy định về quản trị ESG của PVEP (xây mới)", loai:"Phát triển bền vững", ban:"CL&QTRR", trangthai:"chua", u:"", task:"6.1",
       kp:["Xây mới theo nhiệm vụ 6.1 KH 1017 — tích hợp ESG, rủi ro khí hậu và lộ trình chuyển dịch năng lượng vào chiến lược, đầu tư, tài chính, QTRR (hạn Quý IV/2026)."]},
      {so:"", ten:"Quy định mới về công bố thông tin, minh bạch", loai:"Minh bạch – CBTT", ban:"VP", trangthai:"dangsua", u:"", task:"7.1",
       kp:["Thay quy chế CBTT hiện hành — thống nhất danh mục thông tin công bố, quy trình chuẩn hóa, báo cáo thường niên song ngữ tháng 4 hằng năm — nhiệm vụ 7.1 KH 1017 (hạn Quý III/2026)."]},
      {so:"", ten:"Quy trình KTGS & KTNB chuẩn hóa theo chuẩn mực IIA", loai:"Kiểm soát – rủi ro", ban:"KSNB", trangthai:"dangsua", u:"", task:"5.1",
       kp:["Chuẩn hóa và số hóa quy trình kiểm tra giám sát, kiểm toán nội bộ theo chuẩn mực quốc tế IIA — nhiệm vụ 5.1 KH 1017 (hạn Quý III/2026)."]},
      {so:"", ten:"Quy định CNTT, an toàn mạng (sửa đổi — rủi ro số, SCADA)", loai:"Kiểm soát – rủi ro", ban:"CN-CĐS", trangthai:"chua", u:"", task:"4.3",
       kp:["Tích hợp rủi ro số, an ninh mạng, SCADA/công nghệ vận hành vào hồ sơ rủi ro và khung QTRR — nhiệm vụ 4.3 KH 1017 (hạn Quý III/2026)."]},
      {so:"", ten:"Quy chế quản lý Người đại diện (hợp nhất, hoàn thiện)", loai:"Người đại diện", ban:"QTNNL", trangthai:"chua", u:"", task:"8.1",
       kp:["Hợp nhất khung quản trị Người đại diện vốn và Người đại diện tại dự án dầu khí; ma trận năng lực, khung KPI gắn kết quả quản trị — nhiệm vụ 8.1 KH 1017 (hạn Quý IV/2026)."]},
    ],
  };

  /* =========================================================================
     HÀM TÍNH — mọi KPI đều tính live từ dữ liệu trên, KHÔNG hard-code
     ========================================================================= */

  /* Trạng thái nhiệm vụ: done | red (quá hạn) | amber (sắp hạn/nguy cơ) | green */
  function taskStatus(t) {
    if (t.p >= 100) return { cls: "done",  label: "Hoàn thành" };
    if (!t.cont && t.e < NOW_IDX) return { cls: "red", label: "Quá hạn" };
    const left = t.e - NOW_IDX;                       // số tháng còn lại tới hạn
    if (!t.cont && left <= 2 && t.p < 60) return { cls: "amber", label: left <= 1 ? "Sắp đến hạn" : "Nguy cơ chậm" };
    if (t.cont && t.p < 20 && t.s <= NOW_IDX) return { cls: "amber", label: "Cần đẩy nhanh" };
    return { cls: "green", label: t.p > 0 ? "Đang thực hiện" : "Trong kế hoạch" };
  }

  function stats() {
    const s = { total: TASKS.length, done: 0, red: 0, amber: 0, green: 0, decide: 0, avg: 0 };
    TASKS.forEach((t) => {
      s[taskStatus(t).cls]++;
      if ((t.requests || "").trim()) s.decide++;
      s.avg += t.p;
    });
    s.avg = Math.round((s.avg / s.total) * 10) / 10;
    return s;
  }

  /* Thang chấm 5 nấc: 1 Chưa có → 5 Thông lệ tốt. Điểm = Σ w·(score−1)/4 ÷ Σw × 100 */
  const SCORE_LABELS = ["Chưa có","Sơ khai","Đang hoàn thiện","Cơ bản đạt","Thông lệ tốt"];
  const critStatus = (c) => c.score >= 4 ? "dat" : c.score === 3 ? "dangsua" : "hut";
  function scoreLive(pillar) {
    const list = CRITERIA.filter((c) => !pillar || c.pillar === pillar);
    const tw = list.reduce((a, c) => a + c.w, 0);
    if (!tw) return 0;
    const got = list.reduce((a, c) => a + c.w * (c.score - 1) / 4, 0);
    return Math.round((got / tw) * 1000) / 10;
  }
  function maturity(pct) {
    if (pct >= 90) return { n: 5, name: "Dẫn dắt",    hint: "Duy trì và chia sẻ thông lệ tốt trong Tập đoàn." };
    if (pct >= 75) return { n: 4, name: "Chủ động",   hint: "Đóng các tiêu chí đang sửa trọng số cao trước Quý IV/2026." };
    if (pct >= 55) return { n: 3, name: "Chuẩn hóa",  hint: "Ưu tiên ban hành các quy chế nền (ESG, CBTT, QTRR)." };
    if (pct >= 35) return { n: 2, name: "Hình thành", hint: "Tập trung nguồn lực cho nhóm tiêu chí trọng số 3." };
    return            { n: 1, name: "Khởi động",  hint: "Thiết lập đầu mối và kế hoạch chi tiết từng trụ cột." };
  }
  function gaps() {
    return CRITERIA.filter((c) => c.score < 4)
      .map((c) => ({ ...c, status: critStatus(c), prio: c.w * (c.score <= 2 ? 2 : 1) }))
      .sort((a, b) => b.prio - a.prio);
  }
  /* Phả hệ chỉ đạo (PH4) */
  const dirById = (id) => DIRECTIVES.find((d) => d.id === id);
  const dirChildren = (id) => DIRECTIVES.filter((d) => d.parentId === id);

  /* Tra cứu 3 chiều */
  const byId = (arr, id, k = "id") => arr.find((x) => x[k] === id);
  const taskById = (id) => byId(TASKS, id);
  const critById = (id) => byId(CRITERIA, id);
  const evById   = (id) => byId(EVIDENCE, id);
  const evidenceForCriterion = (cid) => EVIDENCE.filter((e) => (e.criteria || []).includes(cid));
  const tasksForCriterion    = (cid) => TASKS.filter((t) => (t.criteria || []).includes(cid));
  const evidenceForTask      = (tid) => EVIDENCE.filter((e) => (e.tasks || []).includes(tid));

  return { META, sp, spO, GMONTHS, NOW_IDX, GROUPS, TASKS, PILLARS, CRITERIA, EVIDENCE, SCORES, DOCS,
           LEVELS, DIRECTIVES, dirById, dirChildren,
           taskStatus, stats, scoreLive, maturity, gaps, SCORE_LABELS, critStatus,
           taskById, critById, evById, evidenceForCriterion, tasksForCriterion, evidenceForTask };
})();
