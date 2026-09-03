# QUY TẮC TẠO VÀ CHỈNH SỬA BÀI TẬP TIẾNG ANH

## 1. Phạm vi áp dụng

Quy tắc này bắt buộc áp dụng cho:

- Bài tập của từng ngày học.
- Đề bài tập tổng hợp theo nhóm bài.
- Bộ đề ôn thi tốt nghiệp THPT và xét tuyển đại học.
- Mọi lần bổ sung, sửa chữa hoặc định dạng lại một bài tập đã có.

Quy ước riêng về lời giải:

- Chỉ các tài liệu nằm trong thư mục `exercises\Bộ đề ôn thi tốt nghiệp THPT và xét tuyển đại học` mới bắt buộc có giải thích chi tiết cho từng câu.
- Mọi bài tập hoặc đề nằm ngoài thư mục trên không bắt buộc có giải thích; mặc định phần đáp án chỉ ghi đáp án đúng hoặc câu trả lời mẫu.
- Căn cứ phân loại là đường dẫn thực tế của tệp, không phải tên đề, độ khó hoặc tiêu đề bên trong tài liệu.

Khi một đề cũ được chỉnh sửa, phải kiểm tra lại toàn bộ đề và phần đáp án theo quy tắc này, không chỉ kiểm tra riêng đoạn vừa sửa.

## 1.1. Bộ công cụ xử lý Word

- Khi tạo, sửa, đọc hoặc kiểm tra DOCX bằng Python, bắt buộc dùng `.venv-word-tools/bin/python` của project.
- Bộ thư viện chuẩn trong môi trường này là `python-docx==1.2.0`, `lxml==6.1.3`, `typing_extensions==4.16.0`; nếu cần cài hoặc kiểm tra package, dùng `.venv-word-tools/bin/python -m pip`.
- Không dùng Python hệ thống để ghi DOCX khi `.venv-word-tools` tồn tại, nhằm tránh khác phiên bản thư viện giữa các project.
- Tệp kiểm tra, bản sao tạm, bản convert PDF và sản phẩm trung gian phải đặt trong thư mục tạm ngoài workspace, trừ khi người dùng yêu cầu lưu lại.
- Sau khi ghi DOCX, phải kiểm tra `unzip -t`, phân tích mọi XML/RELS bằng `lxml`, kiểm tra quan hệ gói và thử mở bằng LibreOffice headless trước khi bàn giao.

## 2. Nguyên tắc chung

1. Nội dung phải đúng với kiến thức người học đã được học tại thời điểm làm bài.
2. Bài tập phải có mục tiêu rõ ràng, độ khó phù hợp và không dùng kiến thức chưa được giới thiệu nếu không có chú thích hỗ trợ.
3. Mỗi câu hỏi chỉ nên kiểm tra một trọng tâm chính, trừ các câu đọc hiểu hoặc vận dụng tổng hợp.
4. Câu hỏi, lựa chọn và đáp án phải thống nhất tuyệt đối về số thứ tự, nội dung và ký hiệu.
5. Không viết thành một khối chữ dài. Phải chia đoạn theo ý nghĩa và tạo khoảng trống hợp lý để người học dễ đọc.
6. Ưu tiên câu văn tự nhiên, đúng ngữ pháp và có ngữ cảnh thực tế; tránh câu đánh đố hoặc phương án nhiễu mơ hồ.

## 3. Cấu trúc bắt buộc của một bài tập hoặc đề

### 3.1. Phần đầu

Phải có các thông tin phù hợp với loại bài:

- Tên bài tập hoặc số đề.
- Phạm vi kiến thức hoặc chủ đề.
- Thời gian làm bài dự kiến.
- Tổng số câu và thang điểm nếu có.
- Hướng dẫn làm bài ngắn gọn.

Không đặt đáp án, gợi ý làm bài hoặc thông tin làm lộ đáp án ở phần đầu.

### 3.2. Phần câu hỏi

- Chia thành từng phần theo dạng bài hoặc kỹ năng.
- Mỗi phần có tiêu đề riêng, ví dụ: `PHẦN 1. NGỮ ÂM (Câu 1–4)`.
- Ngay dưới tiêu đề phải có một câu hướng dẫn cụ thể.
- Đánh số câu liên tục, không trùng và không bỏ số.
- Nếu đổi dạng bài, phải bắt đầu một phần mới.

### 3.3. Phần đáp án

- Luôn đặt ở cuối cùng của tài liệu, sau toàn bộ câu hỏi và phần tự đánh giá nếu có.
- Với đề dài, bắt đầu phần đáp án ở một trang mới.
- Với tài liệu ngoài thư mục `exercises\Bộ đề ôn thi tốt nghiệp THPT và xét tuyển đại học`, dùng tiêu đề `ĐÁP ÁN`.
- Với tài liệu trong thư mục `exercises\Bộ đề ôn thi tốt nghiệp THPT và xét tuyển đại học`, dùng tiêu đề `ĐÁP ÁN VÀ GIẢI THÍCH CHI TIẾT`.
- Chia đáp án theo phần hoặc nhóm kỹ năng tương ứng với đề.
- Không đặt thêm nội dung luyện tập, ghi chú hay câu hỏi mới sau phần đáp án.

## 4. Quy tắc trình bày và xuống dòng

### 4.1. Tiêu đề và đoạn văn

- Tiêu đề chính căn giữa, nổi bật và không viết chung dòng với nội dung khác.
- Tiêu đề phần đứng trên một dòng riêng.
- Hướng dẫn làm bài đứng ở một đoạn riêng ngay dưới tiêu đề phần.
- Mỗi đoạn văn chỉ nên thể hiện một ý hoặc một nhóm ý liên quan.
- Đoạn đọc dài phải được chia thành các đoạn nhỏ theo nội dung; không dồn toàn bài đọc vào một đoạn duy nhất.
- Hội thoại: mỗi lượt lời của một người phải xuống dòng riêng.
- Email/thư: lời chào, nội dung chính và lời kết phải tách thành các đoạn hợp lý.
- Thông báo/tờ rơi: tiêu đề, thông tin chính, thời gian, địa điểm và lời kêu gọi nên tách dòng theo chức năng.

#### 4.1.1. Khoảng cách đoạn bắt buộc

- Không chỉ dùng ký tự xuống dòng để tạo khoảng cách. Phải thiết lập khoảng cách đoạn bằng thuộc tính `Space Before` và `Space After` của đoạn văn.
- Tiêu đề chính: cách đoạn sau từ `12 pt` đến `18 pt`.
- Tiêu đề phần: cách đoạn trước từ `12 pt` đến `16 pt`, cách đoạn sau từ `3 pt` đến `5 pt`.
- Riêng `PHẦN 1` ngay sau phần thông tin đầu đề/hướng dẫn chung: khoảng cách trước chỉ từ `2 pt` đến `4 pt`; không áp dụng mức `12–16 pt` của các phần tiếp theo.
- Câu hướng dẫn dưới tiêu đề phần: cách đoạn sau từ `3 pt` đến `5 pt`.
- Câu hỏi đầu tiên của phần không được cộng thêm khoảng cách lớn ở phía trước; tổng khoảng cách thị giác từ hướng dẫn đến câu đầu tiên chỉ nên từ `4 pt` đến `6 pt`.
- Không để tổ hợp khoảng cách của tiêu đề phần, hướng dẫn và câu hỏi đầu tiên tạo thành một vùng trắng lớn hoặc làm phần câu hỏi bị đẩy sang trang sau.
- Các dòng họ tên, phạm vi kiến thức, thời gian và hướng dẫn chung ở phần đầu chỉ cách nhau từ `2 pt` đến `4 pt`. Sau hướng dẫn chung không được có đoạn trống trước `PHẦN 1`.
- Không chèn ngắt trang, ngắt vùng (`section break`) hoặc nhiều dấu xuống dòng giữa phần đầu đề và `PHẦN 1`.
- Đoạn văn thông thường: cách đoạn sau tối thiểu `6 pt`; các đoạn dài hoặc chuyển ý nên dùng từ `8 pt` đến `10 pt`.
- Không đặt tiêu đề phần sát câu cuối của phần trước. Giữa hai phần phải có khoảng trắng thị giác rõ ràng, tương đương ít nhất `16 pt`.
- Không dùng nhiều dòng trống liên tiếp để thay cho khoảng cách đoạn vì dễ gây sai lệch khi mở trên thiết bị hoặc phần mềm khác.

#### 4.1.2. Giữ đoạn kế tiếp và ngắt trang

- Chỉ dùng thuộc tính `Keep with next` cho một cụm ngắn cần đi cùng nhau: tiêu đề chính với dòng phụ đề, tiêu đề phần với câu hướng dẫn, hoặc câu hỏi với chính các lựa chọn A–D của câu đó.
- Không áp dụng `Keep with next` liên tục cho toàn bộ các câu, toàn bộ lựa chọn, cả một phần bài hoặc cả phần đáp án. Chuỗi này có thể làm Word đẩy `PHẦN 1` sang trang sau và tạo một vùng trắng lớn giữa phần giới thiệu đề và câu hỏi đầu tiên.
- Không được để chuỗi `Keep with next` kéo dài quá cụm tiêu đề phần – hướng dẫn – câu hỏi đầu tiên. Đoạn cuối của cụm phải đặt `Keep with next = Off` để Word có điểm ngắt trang tự nhiên.
- Với tài liệu cũ có câu hỏi và các lựa chọn nằm chung trong một đoạn, đoạn câu hỏi đó bắt buộc đặt `Keep with next = Off`; không được liên kết đoạn này với câu hỏi kế tiếp.
- Lựa chọn cuối cùng của một câu, thường là `D.`, không đặt `Keep with next`; câu hỏi tiếp theo phải bắt đầu theo khoảng cách đoạn đã quy định, không bị nối thành cùng một chuỗi.
- Không chèn ngắt trang thủ công, `Page break before` hoặc giữ lại dấu `lastRenderedPageBreak` ngay trước `PHẦN 1` nếu phần giới thiệu và tiêu đề phần vẫn còn đủ chỗ trên cùng trang. `PHẦN 1` phải nằm ngay sau phần hướng dẫn chung với khoảng cách theo Mục 4.1.1.
- Nếu nội dung thực sự không còn đủ chỗ, chỉ chấp nhận ngắt trang tự nhiên hoặc ngắt trang có chủ đích đã được kiểm tra trực quan; không để ngắt trang do chuỗi `Keep with next` gây ra.
- Sau khi lưu phải cập nhật phân trang và kiểm tra số trang thực tế: nếu dòng `Họ và tên` ở trang 1 nhưng `PHẦN 1` bị chuyển sang trang 2 trong khi cuối trang 1 còn vùng trắng lớn, tài liệu không đạt và phải kiểm tra lại `Keep with next` từ `PHẦN 1` đến các câu tiếp theo.

### 4.2. Câu hỏi và lựa chọn

- Mỗi câu hỏi bắt đầu ở một dòng mới.
- Không để hai câu hỏi trên cùng một dòng.
- Nội dung câu hỏi dùng kiểu chữ thường, không in đậm toàn bộ câu. Chỉ in đậm từ khóa hoặc cụm từ cần nhấn mạnh khi dạng bài thực sự yêu cầu.
- Với câu trắc nghiệm, mỗi lựa chọn `A.`, `B.`, `C.`, `D.` phải nằm trên một dòng riêng.
- Không ghép bốn lựa chọn thành một dòng, kể cả khi lựa chọn ngắn.
- Giữ khoảng cách nhỏ giữa câu hỏi và các lựa chọn; tạo khoảng cách lớn hơn trước câu hỏi tiếp theo.
- Không để câu hỏi ở cuối trang trong khi toàn bộ lựa chọn bị đẩy sang trang sau nếu có thể tránh được.
- Các từ/cụm từ được hỏi phải được đánh dấu nhất quán bằng gạch chân, in đậm hoặc ký hiệu chỗ trống.

Mẫu trình bày:

```text
1. She ______ to school every day.

A. walk
B. walks
C. walked
D. walking
```

Khoảng cách khuyến nghị cho câu trắc nghiệm:

- Câu hỏi thông thường: cách đoạn trước từ `5 pt` đến `7 pt`, cách lựa chọn đầu tiên từ `2 pt` đến `4 pt`.
- Câu hỏi đầu tiên ngay sau hướng dẫn: cách đoạn trước tối đa `4 pt`.
- Giữa các lựa chọn A/B/C/D: cách đoạn sau từ `2 pt` đến `4 pt`.
- Sau lựa chọn cuối cùng: cách câu hỏi tiếp theo từ `10 pt` đến `12 pt`.
- Không để lựa chọn cuối của câu trước dính sát vào câu hỏi tiếp theo.
- Với câu tự luận, dịch, sửa lỗi hoặc sắp xếp từ, mỗi câu cách câu tiếp theo tối thiểu `8 pt`.

### 4.3. Câu sắp xếp và ghép nối

- Mỗi mệnh đề hoặc lượt lời `a.`, `b.`, `c.`, `d.` nằm trên một dòng riêng.
- Các phương án thứ tự `A.`, `B.`, `C.`, `D.` cũng nằm trên các dòng riêng.
- Phân biệt rõ dữ liệu cần sắp xếp với các phương án trả lời.

### 4.4. Định dạng nhất quán

- Dùng một kiểu phông chữ dễ đọc trong toàn tài liệu.
- Cỡ chữ nội dung phải đủ lớn; tiêu đề có cấp độ rõ ràng.
- Màu chữ mặc định của toàn bộ tài liệu phải là màu đen (`#000000`) hoặc `Automatic` hiển thị thành màu đen trên nền trắng.
- Phải đặt lại màu chữ trực tiếp cho toàn bộ nội dung sau khi sao chép, chuyển đổi hoặc áp dụng mẫu; không được chỉ dựa vào màu kế thừa từ `Style`, `Theme`, hyperlink hoặc tài liệu nguồn.
- Câu hỏi, lựa chọn, đoạn đọc, hội thoại, hướng dẫn và đáp án bắt buộc dùng chữ màu đen; không dùng màu xanh đồng loạt cho nội dung.
- Tiêu đề ưu tiên dùng chữ đen và in đậm. Chỉ dùng màu chữ khác khi mẫu tài liệu riêng yêu cầu rõ ràng; màu đó chỉ được áp dụng đúng phạm vi tiêu đề, không lan sang các đoạn bên dưới.
- Khi đặt chữ về màu đen, phải xóa đồng thời các thuộc tính màu theme như `accent`, `hyperlink` và `followed hyperlink` nếu chúng đang được kế thừa ngoài ý muốn.
- Không tô màu nền (`shading`) hoặc bôi màu (`text highlight`) cho phần nội dung, câu hỏi, lựa chọn, đoạn đọc, hội thoại và đáp án.
- Mặc định mọi đoạn nội dung phải có nền trắng hoặc `No Color`, đồng thời thuộc tính `Text Highlight Color` phải là `No Color`.
- Tiêu đề cũng ưu tiên không dùng màu nền. Chỉ được dùng nền cho tiêu đề khi mẫu tài liệu riêng yêu cầu rõ ràng; màu phải rất nhạt, thống nhất và không kéo dài sang các đoạn nội dung bên dưới.
- Khi sao chép nội dung từ tài liệu hoặc nguồn khác, phải xóa định dạng nền và highlight kế thừa trước khi áp dụng kiểu chữ của tài liệu.
- Không dùng màu nền để tạo khoảng cách hoặc phân chia phần; phải dùng khoảng cách đoạn, tiêu đề và ngắt trang.
- Không lạm dụng viết hoa, gạch chân hoặc nhiều màu trong cùng một đoạn.
- Dấu câu phải đặt đúng vị trí; có một khoảng trắng sau dấu phẩy, dấu chấm, dấu chấm phẩy và dấu hai chấm.
- Không có lỗi “Ngắt cầu”; phải kiểm tra và sửa thành “ngắt câu”, xuống dòng đúng ngữ nghĩa.

## 5. Quy tắc biên soạn câu hỏi

### 5.1. Câu trắc nghiệm

- Chỉ có một đáp án đúng rõ ràng, trừ khi hướng dẫn nêu hình thức khác.
- Các phương án phải cùng loại từ, cùng cấu trúc hoặc có hình thức tương xứng khi phù hợp.
- Phương án nhiễu phải hợp lý nhưng sai vì một căn cứ xác định.
- Không để đáp án đúng nổi bật do dài hơn bất thường, khác định dạng hoặc lặp nguyên văn quá rõ.
- Phân bố đáp án A/B/C/D tương đối cân bằng; tránh chuỗi đáp án giống nhau quá dài nếu không cần thiết.

### 5.2. Đọc hiểu

- Văn bản phải có nguồn ngôn ngữ tự nhiên và phù hợp trình độ.
- Câu hỏi phải bao phủ nhiều mức: thông tin chi tiết, ý chính, từ vựng trong ngữ cảnh, tham chiếu, suy luận và mục đích/thái độ khi phù hợp.
- Câu hỏi suy luận phải được thiết kế để đáp án có thể kiểm chứng bằng bằng chứng từ văn bản, không phụ thuộc vào kiến thức ngoài bài.
- Nếu hỏi từ đồng nghĩa/trái nghĩa, nghĩa được chọn phải đúng trong ngữ cảnh cụ thể.

### 5.3. Điền khuyết

- Chỗ trống phải có đủ ngữ cảnh để xác định đáp án.
- Kiểm tra sự hòa hợp chủ ngữ–động từ, thì, giới từ, từ loại, liên từ và logic toàn đoạn.
- Không tạo hai đáp án đều đúng vì thiếu ngữ cảnh.

### 5.4. Câu trả lời mở và bài viết

- Nêu rõ yêu cầu, độ dài, nội dung bắt buộc và tiêu chí đánh giá.
- Phải có bài làm mẫu hoàn chỉnh ở phần đáp án.
- Bài mẫu chỉ là một phương án tham khảo, dùng ngôn ngữ phù hợp với trình độ đã học.
- Kèm dàn ý hoặc các ý chính và thang/tiêu chí chấm khi bài yêu cầu sản phẩm dài.
- Nếu có nhiều cách trả lời đúng, ghi rõ `Câu trả lời tham khảo` và nêu nguyên tắc chấp nhận phương án khác.

### 5.5. Tính khác biệt giữa các đề trong cùng một bộ

- Các đề trong cùng một thư mục phải là những phiên bản độc lập, không chỉ khác số đề, tên nhân vật, địa điểm, ngày tháng hoặc một vài từ trong ngữ cảnh.
- Có thể dùng chung cấu trúc phần, dạng bài, hướng dẫn và tiêu chí chấm; không được dùng chung nguyên văn câu hỏi, đoạn đọc, hội thoại, email, tờ rơi, phương án nhiễu hoặc thứ tự đáp án.
- Khi tạo đề biến thể, tối thiểu `70%` đơn vị câu hỏi của mỗi cặp đề phải khác nhau. Một đơn vị câu hỏi gồm câu dẫn, đoạn văn/hội thoại liên quan, các lựa chọn và yêu cầu trả lời.
- Không để quá `2/10` phần của hai đề có nội dung câu hỏi giống hoàn toàn. Những phần bắt buộc dùng chung mẫu phải được thay đổi câu hỏi, dữ liệu, ngữ cảnh và phương án nhiễu.
- Việc chỉ đổi tên riêng, con số, chủ đề bề mặt hoặc đảo vị trí lựa chọn không được tính là tạo nội dung mới nếu căn cứ làm bài và đáp án vẫn giữ nguyên.
- Với phần đọc hiểu và viết, mỗi đề phải có văn bản/ngữ cảnh và yêu cầu riêng; không dùng lại cùng một đoạn đọc rồi chỉ thay tên chủ đề.
- Trước khi lưu, phải so sánh chéo toàn bộ các đề trong cùng thư mục sau khi đã bỏ tiêu đề, thông tin điểm/thời gian, hướng dẫn chung và khoảng trắng. Nếu vượt ngưỡng trùng nêu trên, phải biên soạn lại trước khi hoàn tất.

## 6. Chuẩn bắt buộc cho phần đáp án

### 6.0. Phân loại theo đường dẫn

Trước khi tạo hoặc sửa phần đáp án, phải xác định tệp thuộc một trong hai nhóm:

1. **Nhóm đáp án ngắn:** mọi tệp nằm ngoài `exercises\Bộ đề ôn thi tốt nghiệp THPT và xét tuyển đại học`.
2. **Nhóm đáp án có giải thích:** mọi tệp nằm trong `exercises\Bộ đề ôn thi tốt nghiệp THPT và xét tuyển đại học`, kể cả các thư mục con của thư mục này.

Không được áp dụng yêu cầu giải thích chi tiết của nhóm 2 cho nhóm 1 chỉ vì đề dài hoặc khó.

### 6.1. Nội dung tối thiểu

Mỗi câu phải có:

1. Số câu.
2. Đáp án đúng hoặc câu trả lời mẫu.

Với nhóm đáp án ngắn, phần đáp án chỉ cần ghi đáp án/câu trả lời mẫu. Mặc định không ghi `Giải thích:`, bằng chứng, suy luận, bản dịch, phân tích phương án nhiễu hoặc lời giải bổ sung.

Với nhóm đáp án có giải thích, ngoài số câu và đáp án đúng, mỗi câu phải có dòng `Giải thích:` riêng theo Mục 6.4.

Với câu trả lời mở, phải có bài viết mẫu hoàn chỉnh và có thể ghi thêm tiêu chí chấm hoặc nguyên tắc chấp nhận phương án khác nếu đề yêu cầu.

### 6.2. Cách ghi theo dạng bài

- Câu trắc nghiệm: ghi số câu, chữ cái đáp án và nội dung đáp án khi cần, ví dụ `Câu 1: B. walks`.
- Câu điền khuyết: ghi trực tiếp từ/cụm từ cần điền, ví dụ `Câu 2: at`.
- Câu sắp xếp: ghi đầy đủ câu hoặc thứ tự đúng, không chỉ ghi chữ cái phương án.
- Câu dịch, sửa lỗi và tự luận: ghi câu trả lời mẫu hoàn chỉnh.
- Nếu có nhiều cách trả lời đúng, ghi `Câu trả lời tham khảo` và nêu nguyên tắc chấp nhận phương án khác, không kèm lời giải.

### 6.3. Trình bày phần đáp án

- Tiêu đề phần đáp án nằm ở cuối tài liệu; dùng `ĐÁP ÁN` cho nhóm đáp án ngắn và `ĐÁP ÁN VÀ GIẢI THÍCH CHI TIẾT` cho nhóm đáp án có giải thích. Với đề dài, bắt đầu ở trang mới.
- Tiêu đề nhóm đáp án phải tương ứng với phần câu hỏi và nằm trên một dòng riêng, ví dụ: `PHẦN 3` hoặc `NHÓM 3. ĐỌC HIỂU`.
- Mỗi câu trả lời bắt đầu ở một đoạn mới và chỉ ghi số câu cùng đáp án/câu trả lời mẫu.
- Không ghép nhiều đáp án của nhiều câu vào cùng một đoạn văn.
- Giữa các dòng đáp án dùng khoảng cách đoạn nhất quán, đủ để phân biệt từng câu nhưng không tạo vùng trắng lớn; ưu tiên `Space After` từ `6 pt` đến `10 pt`.
- Các bài mẫu, tiêu chí chấm và ghi chú chấp nhận phương án khác phải nằm ở các đoạn riêng sau nhãn tương ứng.
- Toàn bộ phần đáp án dùng nền trắng hoặc `No Color`, không dùng highlight để đánh dấu đáp án đúng.
- Không để tiêu đề nhóm ở cuối trang trong khi dòng đáp án đầu tiên bị chuyển hoàn toàn sang trang sau nếu có thể tránh được.

### 6.4. Yêu cầu riêng cho bộ đề ôn thi tốt nghiệp THPT và xét tuyển đại học

Mục này chỉ áp dụng cho tài liệu trong `exercises\Bộ đề ôn thi tốt nghiệp THPT và xét tuyển đại học`:

- Mỗi câu trình bày thành một khối gồm dòng `Câu ...: Đáp án ...` và một dòng riêng bắt đầu bằng `Giải thích:`.
- Câu dễ vẫn phải nêu quy tắc, từ khóa hoặc bằng chứng trực tiếp; không dùng lời giải chung chung như “đúng ngữ pháp” hoặc “phù hợp nhất”.
- Câu đọc hiểu phải chỉ rõ vị trí bằng chứng trong đoạn; câu suy luận phải tách được bằng chứng, suy luận và kết luận.
- Câu khó hoặc có phương án gần nghĩa phải phân tích phương án nhiễu khi cần; mỗi phương án phân tích nằm trên một dòng riêng.
- Câu sắp xếp phải giải thích mạch liên kết; câu phát âm/trọng âm phải chỉ rõ âm hoặc vị trí trọng âm khác biệt.
- Đáp án và dòng `Giải thích:` không được ghép vào cùng một dòng.
- Sau mỗi khối đáp án–giải thích dùng khoảng cách đoạn nhất quán từ `10 pt` đến `12 pt`.

## 7. Mức độ và tính phù hợp

- Bài tập ngày học: tập trung vào mục tiêu của ngày, có một phần ôn kiến thức cũ vừa đủ.
- Đề tổng hợp: bao phủ cân đối các bài trong nhóm, không thiên lệch quá mức về một bài.
- Đề thi thử: bám sát cấu trúc thi đang áp dụng trong bộ tài liệu và tăng độ khó có kiểm soát.
- Độ khó tăng bằng chiều sâu của ngữ cảnh, paraphrase, suy luận và độ gần của phương án nhiễu; không tăng bằng câu chữ tối nghĩa.
- Từ vựng vượt trình độ chỉ xuất hiện khi có thể suy ra từ ngữ cảnh hoặc có chú thích phù hợp.

## 8. Kiểm tra bắt buộc trước khi lưu

### 8.1. Kiểm tra nội dung

- [ ] Nội dung đúng phạm vi bài học/lộ trình.
- [ ] Mỗi câu có một đáp án xác định.
- [ ] Không có câu trùng lặp hoặc mâu thuẫn.
- [ ] Câu tiếng Anh đúng chính tả và ngữ pháp.
- [ ] Đáp án thực tế khớp với câu hỏi và phương án đúng.
- [ ] Câu trả lời mở có bài mẫu.
- [ ] Các đề trong cùng một bộ đã được so sánh chéo; không có cặp đề vượt ngưỡng trùng nội dung ở Mục 5.5.
- [ ] Những phần khác nhau giữa các đề là khác về câu hỏi, ngữ cảnh và phương án nhiễu, không chỉ khác tên, số liệu hoặc thứ tự đáp án.

### 8.2. Kiểm tra cấu trúc

- [ ] Số câu liên tục, không thiếu và không trùng.
- [ ] Mỗi câu trắc nghiệm có đủ lựa chọn theo yêu cầu.
- [ ] Tiêu đề phần và phạm vi số câu chính xác.
- [ ] Không có nội dung câu hỏi mới sau phần đáp án.
- [ ] Phần đáp án có đủ số lượng câu của đề.

### 8.3. Kiểm tra trình bày

- [ ] Mỗi câu hỏi ở một dòng/đoạn riêng.
- [ ] Mỗi lựa chọn A/B/C/D ở một dòng riêng.
- [ ] Đoạn đọc, hội thoại, email và thông báo đã xuống dòng theo ý nghĩa.
- [ ] Khoảng cách đoạn nhất quán, không có khối chữ quá dày.
- [ ] Tiêu đề phần có khoảng cách trước/sau rõ ràng và không dính vào phần trước hoặc câu hỏi đầu tiên.
- [ ] Câu hỏi, các lựa chọn và câu hỏi kế tiếp có khoảng cách đúng theo Mục 4.2; không có hai khối câu dính sát nhau.
- [ ] Tiêu đề phần, hướng dẫn và câu hỏi đầu tiên nằm thành một cụm gọn; không có vùng trắng quá lớn làm kéo dài số trang.
- [ ] `PHẦN 1` nằm sát hợp lý với hướng dẫn chung; khoảng cách trước không vượt `4 pt` và không có đoạn trống/ngắt trang ở giữa.
- [ ] `PHẦN 1` không bị đẩy sang trang sau bởi `Keep with next`, `Page break before` hoặc dấu ngắt trang cũ khi phần giới thiệu vẫn còn đủ chỗ.
- [ ] `Keep with next` chỉ dùng cho các cụm ngắn; không nối liên tục toàn bộ phần câu hỏi hoặc phần đáp án.
- [ ] Đã cập nhật phân trang sau khi lưu; dòng `Họ và tên` và `PHẦN 1` không bị tách sang hai trang do chuỗi `Keep with next` khi trang đầu vẫn còn đủ chỗ.
- [ ] Với câu hỏi có lựa chọn nằm chung một đoạn trong tài liệu cũ, `Keep with next` của đoạn câu hỏi đã được tắt.
- [ ] Câu hỏi không bị in đậm toàn bộ; chỉ từ khóa cần thiết mới được nhấn mạnh.
- [ ] Đáp án bắt đầu ở cuối tài liệu, ưu tiên sang trang mới với đề dài.
- [ ] Đã xác định đúng nhóm đáp án dựa trên đường dẫn tệp theo Mục 6.0.
- [ ] Với tệp ngoài thư mục bộ đề ôn thi: mỗi đáp án bắt đầu ở một dòng/đoạn riêng; phần đáp án ngắn không bị bổ sung lời giải không cần thiết.
- [ ] Với tệp trong thư mục bộ đề ôn thi: mỗi câu có đủ dòng đáp án và dòng `Giải thích:` riêng theo Mục 6.4.
- [ ] Các bài mẫu, tiêu chí chấm và ghi chú chấp nhận phương án khác được tách thành các đoạn riêng khi có yêu cầu.
- [ ] Khoảng cách giữa các dòng đáp án nhất quán và đủ để phân biệt từng câu.
- [ ] Không còn màu nền hoặc highlight trong nội dung, câu hỏi, lựa chọn, đoạn đọc và đáp án.
- [ ] Thuộc tính nền của nội dung là trắng/`No Color`; `Text Highlight Color` là `No Color`.
- [ ] Toàn bộ chữ nội dung hiển thị màu đen; không còn màu xanh hoặc màu theme kế thừa ngoài ý muốn.
- [ ] Đã kiểm tra màu chữ ở cả cấp ký tự, đoạn, `Style` và `Theme`; câu hỏi, lựa chọn và đáp án đều là `#000000` hoặc `Automatic` màu đen.
- [ ] Nếu có tạo/sửa DOCX bằng Python, đã dùng `.venv-word-tools/bin/python` của project.
- [ ] `unzip -t` không báo lỗi; mọi XML/RELS trong gói DOCX phân tích được bằng `lxml`.
- [ ] Các quan hệ trong `.rels`, `[Content_Types].xml`, liên kết ảnh và tài nguyên nhúng còn hợp lệ.
- [ ] LibreOffice headless mở/convert được bản DOCX kiểm tra từ thư mục tạm ngoài workspace.
- [ ] Không có ký tự lỗi, dòng trống thừa hoặc tiêu đề nằm sai vị trí.
- [ ] Tệp mở được bình thường và không hỏng định dạng.

## 9. Quy trình khi sửa một bài tập đã có

1. Đọc toàn bộ đề và đáp án hiện tại.
2. Xác định dạng bài, phạm vi kiến thức và trình độ mục tiêu.
3. Sửa nội dung được yêu cầu.
4. Cập nhật đồng thời số câu, tiêu đề phần, lựa chọn và đáp án liên quan.
5. Định dạng lại đoạn bị ảnh hưởng theo quy tắc xuống dòng.
6. Chạy lại toàn bộ danh sách kiểm tra ở Mục 8.
7. So sánh chéo các đề cùng bộ theo Mục 5.5; nếu phát hiện trùng vượt ngưỡng, phải sửa nội dung trước khi lưu.
8. Bảo đảm lần sửa không làm thay đổi ngoài ý muốn các phần khác của tài liệu.

## 10. Quy ước ưu tiên

Khi yêu cầu riêng của một bài khác với tài liệu này, thực hiện yêu cầu riêng về nội dung nhưng vẫn giữ các chuẩn bắt buộc về tính dễ đọc, xuống dòng, tính đầy đủ và độ chính xác của đáp án. Nếu yêu cầu riêng làm phát sinh mâu thuẫn hoặc khiến câu hỏi có nhiều đáp án đúng, phải điều chỉnh câu hỏi hoặc nêu rõ vấn đề trước khi hoàn tất.
