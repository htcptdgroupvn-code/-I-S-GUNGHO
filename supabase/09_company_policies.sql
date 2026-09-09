-- Bảng lưu "chính sách" từng công ty: danh sách sản phẩm/dịch vụ + ghi chú tỷ lệ,
-- mức thưởng hoa hồng — do Admin tự thêm/sửa/xoá trong trang "Chính sách công ty".
--
-- LƯU Ý QUAN TRỌNG: bảng này chỉ là tài liệu tham khảo/ghi chú chính sách, KHÔNG
-- tự động thay đổi công thức tính hoa hồng thật của app (các công thức đó đang
-- được viết cứng trong code theo từng loại sản phẩm, vì mỗi loại có quy tắc tính
-- khác nhau và đã chạy đúng từ trước). Nếu sau này muốn con số ở đây tự động áp
-- dụng vào cách tính hoa hồng thật, cần làm thêm 1 bước riêng để nối 2 phần này.
create table if not exists company_policies (
  id uuid primary key default gen_random_uuid(),
  company_name text not null,
  product_name text not null,
  commission_note text,
  updated_by_name text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

comment on table company_policies is
  'Ghi chú chính sách/tỷ lệ thưởng hoa hồng theo từng công ty và sản phẩm — Admin tự quản lý trong app. Chỉ mang tính tham khảo, không tự động đổi công thức tính hoa hồng thật.';
