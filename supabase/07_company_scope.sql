-- Cách ly dữ liệu theo công ty (5 công ty của Gungho)
-- Chạy 1 lần trong Supabase → SQL Editor. Không xoá/sửa dữ liệu cũ.

alter table employees
  add column if not exists visible_companies text[];

comment on column employees.visible_companies is
  'Danh sách công ty (đúng tên trong COMPANIES ở App.jsx) mà tài khoản này được phép xem. '
  'NULL hoặc rỗng = mặc định chỉ xem đúng công ty của chi nhánh (store) mình đang gắn. '
  'Gán mảng nhiều công ty cho 1 tài khoản (VD: tài khoản của bạn) để "liên kết" xem nhiều công ty cùng lúc.';

-- Ví dụ gán cho 1 tài khoản admin xem được cả 5 công ty (thay 'admin' bằng đúng employee_code của bạn):
-- update employees set visible_companies = array[
--   'I. Công ty Cổ phần Thương mại I - Khối xe máy',
--   'II. Công ty Cổ phần Thương mại I - Khối ô tô',
--   'III. HTC',
--   'IV. VYC',
--   'V. Vật tư nông nghiệp'
-- ] where employee_code = 'admin';
