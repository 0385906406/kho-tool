// Danh sách tool trong Kho Tool — dùng chung cho 2 bản web:
//   node build-tools.js   -> tools.js (bản trên claude.ai, file nhúng base64)
//   node build-public.js  -> public/  (bản web độc lập, link tải thật, không cần đăng nhập)
// Thêm tool mới: thêm 1 phần tử vào TOOLS (file nguồn đặt trong thư mục web/), rồi chạy lại 2 lệnh trên.
module.exports = [
  {
    id: 'doisoat',
    name: 'Tool Đối Soát NCC ↔ MB51',
    code: 'ToolDoiSoat',
    version: '2026.10.20',
    platform: 'WPS Spreadsheets · macro JS',
    summary: 'So số lượng NCC đã giao với GR trên MB51 theo từng SA/PO, kiểm tra bộ cầu theo xe và chỉ ra chuyến, biên bản bị lệch.',
    features: [
      'Mỗi SA/PO một dòng, NCC đã giao và GR tính bằng công thức SUMIFS',
      'Kiểm tra bộ cầu: 1 xe = 4 mã cầu, 1 chuyến = 4 xe = 16 cái',
      'Tổng hợp theo dòng xe: VF5, VF6, VF7, VF7 Pro, VF8, VF9, Limo',
      'Soi lệch theo khoảng biên bản (BB) dựa trên tờ đầu GR',
      'Sheet Cảnh báo cho dòng dữ liệu thiếu PO, thiếu mã, số lượng sai',
      'Tự nhận cột theo tiêu đề, 30.000 dòng chạy khoảng 5 giây',
    ],
    steps: [
      'Giải nén gói .zip ra một thư mục.',
      'Mở ToolDoiSoat.xlsm bằng WPS Office, chọn Bật macro khi được hỏi.',
      'Trên sheet Control, bấm Chọn file NCC và Chọn file MB51.',
      'Bấm CHECK CHÊNH LỆCH và xem sheet Tổng hợp.',
      'Bấm Lưu kết quả (.xlsx) để giữ kết quả của ngày.',
    ],
    sample: {
      note: 'Ví dụ từ 2 file mẫu trong gói (số liệu giả lập)',
      head: ['SA/PO', 'Dòng xe', 'NCC', 'GR', 'Gap', 'Trạng thái', 'Bộ cầu'],
      rows: [
        ['5500012301', 'VF5', 160, 160, 0, 'ok|Khớp', 'ok|Đủ bộ'],
        ['5500012302', 'VF6', 160, 157, 3, 'bad|NCC thừa / GR thiếu', 'ok|Đủ bộ'],
        ['5500012303', 'VF7', 156, 156, 0, 'ok|Khớp', 'warn|Lệch bộ, 1 chuyến lỗi'],
        ['5500012310', 'VF7 Pro', 160, 0, 160, 'warn|Chưa GR', 'ok|Đủ bộ'],
        ['5500012314', 'Limo', 156, 155, 1, 'bad|NCC thừa / GR thiếu', 'warn|1 chuyến 12/16'],
      ],
    },
    // file đầu tiên = nút tải chính; xlsmOnly = chỉ có ở bản web độc lập (claude.ai không cho tải .xlsm)
    files: [
      { label: 'Bộ đầy đủ', name: 'ToolDoiSoat_2026-10-20.zip', src: 'ToolDoiSoat_2026-10-20.zip', kind: 'zip' },
      { label: 'Chỉ file tool', name: 'ToolDoiSoat.xlsm', src: 'dist/ToolDoiSoat.xlsm', kind: 'xlsm', xlsmOnly: true },
      { label: 'File mẫu NCC', name: 'NCC_mau.xlsx', src: 'dist/NCC_mau.xlsx', kind: 'xlsx' },
      { label: 'File mẫu MB51', name: 'MB51_mau.xlsx', src: 'dist/MB51_mau.xlsx', kind: 'xlsx' },
      { label: 'Mã nguồn JS', name: 'DoiSoat_WPS_JS.txt', src: 'dist/DoiSoat_WPS_JS.txt', kind: 'txt' },
    ],
    changes: [
      'Kiểm tra bộ cầu 4 mã / xe và 16 cái / chuyến',
      'Cột và bảng tổng theo dòng xe VinFast',
      'Sheet chi tiết theo từng SA/PO, liệt kê material bên trong',
      'Thông báo kết quả gọn trong 4 dòng',
    ],
  },
];
