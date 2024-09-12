# Apply-Test
Chạy file bot.py
Nếu không chạy được thì hãy thử cài máy ảo bằng cmd chạy pip package 
pip install virtualenv  
python -m venv selenium_python
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\selenium_python\Scripts\Activate
Đồng thời cài các thư viện sau
pip install selenium
pip install pandas
pip install openpyxl

***Lưu ý: Sửa đường dẫn s = Service("D:/Test/chromedriver-win64/chromedriver.exe") thành đường dẫn ổ đĩa hiện tại
***Lưu ý: Bản chrome driver có thể phải tải lại theo phiên bản chrome bạn đang dùng hiện tại. Bản đang dùng là 128