-- Bảng lưu phân quyền: vai trò nào được thấy mục (tab) nào trong app.
-- Admin tự bật/tắt trong trang "Phân quyền". Nếu 1 (vai trò, mục) chưa có dòng
-- nào ở đây, app tự dùng cấu hình mặc định đang chạy (an toàn, không gì đổi
-- cho tới khi Admin chủ động bật/tắt).
create table if not exists role_permissions (
  role text not null,
  tab_key text not null,
  enabled boolean not null,
  updated_by_name text,
  updated_at timestamptz not null default now(),
  primary key (role, tab_key)
);

comment on table role_permissions is
  'Phân quyền: vai trò nào được thấy mục nào trong app. Admin chỉnh trong trang Phân quyền. Không có dòng = dùng mặc định gốc của hệ thống.';
