# Skill hồ sơ hoàn công

Skill cho Claude: lập hồ sơ hoàn công / hồ sơ quản lý chất lượng công trình cầu đường và bảo
dưỡng thường xuyên đường bộ từ **dự toán + bản vẽ**, theo **file mẫu** của người dùng.

- Skill nằm ở `.claude/skills/ho-so-hoan-cong/` - mở một phiên Claude Code trên kho này là dùng được.
- Hướng dẫn cho Claude: `.claude/skills/ho-so-hoan-cong/SKILL.md`
- Quy ước file mẫu: `.claude/skills/ho-so-hoan-cong/references/quy-uoc-file-mau.md`
- File mẫu: `.claude/skills/ho-so-hoan-cong/assets/`
  - `file-mau-minh-hoa.xlsx` - ví dụ thể hiện quy ước, không phải biểu mẫu chính thức.
  - `file-mau.xlsx` - file mẫu thật (khi có sẽ được ưu tiên dùng).

## Cách dùng

Gửi cho Claude file dự toán (Excel) và bản vẽ (PDF), yêu cầu "lập hồ sơ hoàn công". Kết quả là
một file Excel, mỗi biên bản/phiếu một sheet, kèm danh mục. Ô chưa có số liệu hiện trường được
để trống và tô vàng.

## Mở rộng

Thêm biểu mẫu = thêm một sheet vào file mẫu và một dòng vào sheet `TRINH_TU`. Không cần sửa code.
