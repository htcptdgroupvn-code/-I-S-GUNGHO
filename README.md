# Gungho Sales

Ứng dụng theo dõi doanh thu & đơn hàng Gungho — kết nối database thật qua Supabase.

## Thông tin đã cấu hình sẵn (trong file `.env`)

- Project URL: `https://ufjjfrxalrftjricaxrc.supabase.co`
- Publishable key: đã điền sẵn

## Tài khoản đăng nhập demo ban đầu (mật khẩu chung: `123456`)

| Mã nhân viên | Vai trò |
|---|---|
| ds1 | Đại sứ Gungho |
| ds2 | Đại sứ Gungho |
| xl1 | Nhân viên xử lý - chăm sóc |
| cht1 | Cửa hàng trưởng |
| kt1 | Kế toán |

Bạn có thể thêm/sửa tài khoản trực tiếp trong Supabase: **Table Editor → employees**.

## Chạy thử trên máy tính (tùy chọn, không bắt buộc)

```
npm install
npm run dev
```

## Đưa lên mạng thật (Vercel) — không cần biết code

1. Vào https://vercel.com, đăng nhập bằng tài khoản GitHub (tạo tài khoản GitHub miễn phí trước nếu chưa có)
2. Tải toàn bộ thư mục này lên 1 repository GitHub mới (kéo thả file qua giao diện web GitHub, hoặc dùng GitHub Desktop nếu quen)
3. Trong Vercel, bấm **"Add New" → "Project"** → chọn repository vừa tạo
4. Ở bước cấu hình, mở mục **"Environment Variables"**, thêm đúng 2 dòng:
   - `VITE_SUPABASE_URL` = `https://ufjjfrxalrftjricaxrc.supabase.co`
   - `VITE_SUPABASE_ANON_KEY` = `sb_publishable_Lpyf_LcLis5QBI0FCWb-0A_T3YU53mW`
5. Bấm **"Deploy"** — đợi khoảng 1 phút
6. Xong! Vercel sẽ cho bạn 1 đường link thật (dạng `gungho-sales-xxxx.vercel.app`) — gửi link này cho nhân viên để test diện rộng

## Lưu ý bảo mật (đọc trước khi dùng thật)

- Mật khẩu nhân viên hiện lưu dạng chữ thường (chưa mã hóa) — phù hợp cho giai đoạn test nội bộ, **chưa nên dùng cho dữ liệu khách hàng nhạy cảm lâu dài**
- Ai có đường link đều mở được trang đăng nhập (nhưng không đăng nhập được nếu không có mã + mật khẩu đúng)
- Khi cần lên bản chính thức, nên nâng cấp sang hệ thống đăng nhập chuẩn (Supabase Auth) và thắt chặt lại Row Level Security theo từng vai trò
