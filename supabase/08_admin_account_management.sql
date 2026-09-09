-- Quản lý tài khoản đầy đủ cho Admin (giống app Công nợ HTC): tạo mới, sửa,
-- khoá/mở khoá, xoá — tất cả đều yêu cầu xác nhận đúng mật khẩu của chính
-- Admin đang thao tác (giống cách admin_reset_password đã làm).
-- Chạy 1 lần trong Supabase → SQL Editor. Không xoá/sửa dữ liệu cũ.

alter table employees
  add column if not exists account_disabled boolean not null default false;

comment on column employees.account_disabled is
  'true = tài khoản bị Admin khoá thủ công, không đăng nhập được cho tới khi mở khoá lại.';

-- ---------------------------------------------------------------------------
-- Đăng nhập: chặn luôn tài khoản đã bị khoá thủ công (gộp chung cờ "locked"
-- với lý do khoá vì quá hạn đổi mật khẩu — người dùng chỉ cần biết là bị khoá).
-- ---------------------------------------------------------------------------
DROP FUNCTION IF EXISTS verify_employee_login(text, text);

CREATE OR REPLACE FUNCTION public.verify_employee_login(p_code text, p_password text)
 RETURNS TABLE(id uuid, employee_code text, name text, role text, store text, "position" text, phone text, must_change_password boolean, password_change_deadline timestamp with time zone, locked boolean, visible_companies text[])
 LANGUAGE sql
 SECURITY DEFINER
 SET search_path TO 'public', 'extensions'
AS $function$
  select
    e.id, e.employee_code, e.name, e.role, e.store, e."position", e.phone,
    e.must_change_password, e.password_change_deadline,
    (
      e.account_disabled
      or (e.must_change_password and e.password_change_deadline is not null and now() > e.password_change_deadline)
    ) as locked,
    e.visible_companies
  from employees e
  where e.employee_code = p_code
    and e.password = crypt(p_password, e.password)
  limit 1;
$function$;

-- ---------------------------------------------------------------------------
-- Tạo tài khoản mới — mật khẩu mặc định "123456", bắt đổi ngay lần đăng nhập đầu.
-- Trả về id tài khoản mới nếu thành công; báo lỗi rõ nếu trùng mã nhân viên
-- hoặc sai mật khẩu Admin.
-- ---------------------------------------------------------------------------
DROP FUNCTION IF EXISTS admin_create_employee(uuid, text, text, text, text, text, text, text);

CREATE OR REPLACE FUNCTION public.admin_create_employee(
  p_admin_id uuid, p_admin_password text,
  p_employee_code text, p_name text, p_role text, p_store text, p_position text, p_phone text
)
 RETURNS uuid
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'extensions'
AS $function$
DECLARE
  v_new_id uuid;
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM employees WHERE id = p_admin_id AND role = 'admin' AND password = crypt(p_admin_password, password)
  ) THEN
    RAISE EXCEPTION 'admin_password_incorrect';
  END IF;

  IF EXISTS (SELECT 1 FROM employees WHERE employee_code = p_employee_code) THEN
    RAISE EXCEPTION 'employee_code_taken';
  END IF;

  INSERT INTO employees (employee_code, name, role, store, "position", phone, password, must_change_password)
  VALUES (p_employee_code, p_name, p_role, p_store, p_position, p_phone, crypt('123456', gen_salt('bf')), true)
  RETURNING id INTO v_new_id;

  RETURN v_new_id;
END;
$function$;

-- ---------------------------------------------------------------------------
-- Sửa thông tin tài khoản (tên, vai trò, store, chức vụ, SĐT, danh sách công
-- ty được xem) — không đụng tới mật khẩu.
-- ---------------------------------------------------------------------------
DROP FUNCTION IF EXISTS admin_update_employee(uuid, text, uuid, text, text, text, text, text, text[]);

CREATE OR REPLACE FUNCTION public.admin_update_employee(
  p_admin_id uuid, p_admin_password text, p_target_employee_id uuid,
  p_name text, p_role text, p_store text, p_position text, p_phone text, p_visible_companies text[]
)
 RETURNS boolean
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'extensions'
AS $function$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM employees WHERE id = p_admin_id AND role = 'admin' AND password = crypt(p_admin_password, password)
  ) THEN
    RAISE EXCEPTION 'admin_password_incorrect';
  END IF;

  UPDATE employees SET
    name = p_name, role = p_role, store = p_store, "position" = p_position, phone = p_phone,
    visible_companies = CASE WHEN p_visible_companies IS NOT NULL AND array_length(p_visible_companies, 1) > 0
                              THEN p_visible_companies ELSE NULL END
  WHERE id = p_target_employee_id;

  RETURN true;
END;
$function$;

-- ---------------------------------------------------------------------------
-- Khoá / Mở khoá tài khoản (không cho đăng nhập tạm thời, không xoá dữ liệu).
-- ---------------------------------------------------------------------------
DROP FUNCTION IF EXISTS admin_set_employee_status(uuid, text, uuid, boolean);

CREATE OR REPLACE FUNCTION public.admin_set_employee_status(
  p_admin_id uuid, p_admin_password text, p_target_employee_id uuid, p_disabled boolean
)
 RETURNS boolean
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'extensions'
AS $function$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM employees WHERE id = p_admin_id AND role = 'admin' AND password = crypt(p_admin_password, password)
  ) THEN
    RAISE EXCEPTION 'admin_password_incorrect';
  END IF;

  UPDATE employees SET account_disabled = p_disabled WHERE id = p_target_employee_id;
  RETURN true;
END;
$function$;

-- ---------------------------------------------------------------------------
-- Xoá tài khoản vĩnh viễn. Không cho xoá nếu tài khoản đã gắn với đơn hàng/
-- khách hàng (tránh mất dấu vết dữ liệu cũ) — báo lỗi rõ để dùng nút Khoá thay.
-- ---------------------------------------------------------------------------
DROP FUNCTION IF EXISTS admin_delete_employee(uuid, text, uuid);

CREATE OR REPLACE FUNCTION public.admin_delete_employee(
  p_admin_id uuid, p_admin_password text, p_target_employee_id uuid
)
 RETURNS boolean
 LANGUAGE plpgsql
 SECURITY DEFINER
 SET search_path TO 'public', 'extensions'
AS $function$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM employees WHERE id = p_admin_id AND role = 'admin' AND password = crypt(p_admin_password, password)
  ) THEN
    RAISE EXCEPTION 'admin_password_incorrect';
  END IF;

  IF EXISTS (SELECT 1 FROM orders WHERE created_by = p_target_employee_id OR assigned_handler = p_target_employee_id) THEN
    RAISE EXCEPTION 'employee_has_orders';
  END IF;
  IF EXISTS (SELECT 1 FROM customers WHERE created_by = p_target_employee_id) THEN
    RAISE EXCEPTION 'employee_has_customers';
  END IF;

  DELETE FROM employees WHERE id = p_target_employee_id;
  RETURN true;
END;
$function$;
