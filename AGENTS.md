# Chỉ dẫn làm việc trong dự án

Khi tạo mới hoặc chỉnh sửa bất kỳ tài liệu nào trong thư mục `exercises`, bắt buộc đọc và tuân thủ toàn bộ tệp `QUY_TAC_TRINH_BAY_BAI_TAP.md` tại thư mục gốc.

Các yêu cầu cốt lõi không được bỏ qua:

- Nội dung phải phù hợp với `lotrinh.txt` và các bài học liên quan trong `lessions`.
- Câu hỏi, đoạn văn, hội thoại và lựa chọn phải được ngắt câu, chia đoạn, xuống dòng rõ ràng.
- Mỗi lựa chọn A/B/C/D nằm trên một dòng riêng.
- Phần đáp án phải nằm cuối tài liệu và đủ cho tất cả câu hỏi.
- Mọi câu đều cần căn cứ giải thích; câu khó phải có lời giải chi tiết, bằng chứng và phân tích phương án nhiễu khi cần.
- Trong phần đáp án, dòng ghi đáp án và dòng `Giải thích:` phải tách riêng; không dồn đáp án và lời giải vào một dòng dài.
- Lời giải dài phải chia đoạn theo bằng chứng, suy luận và kết luận; mỗi phương án nhiễu cần phân tích phải nằm trên một dòng riêng.
- Câu trả lời mở phải có bài viết hoặc câu trả lời mẫu.
- Trước khi hoàn tất phải thực hiện danh sách kiểm tra nội dung, cấu trúc và trình bày trong tài liệu quy chuẩn.

## Bộ công cụ xử lý Word

- Khi tạo, sửa, đọc hoặc kiểm tra tệp DOCX trong workspace này, bắt buộc dùng Python của project tại `.venv-word-tools/bin/python`; không dùng Python hệ thống nếu virtualenv này tồn tại.
- Bộ thư viện chuẩn đã cài trong `.venv-word-tools`: `python-docx==1.2.0`, `lxml==6.1.3`, `typing_extensions==4.16.0`.
- Nếu cần dùng `pip`, phải gọi qua `.venv-word-tools/bin/python -m pip`.
- Tệp kiểm tra, bản sao tạm và sản phẩm trung gian phải đặt trong thư mục tạm ngoài workspace, trừ khi người dùng yêu cầu lưu lại.
- Sau khi ghi DOCX, phải kiểm tra ZIP, XML/RELS và khả năng mở bằng LibreOffice headless trước khi bàn giao.
