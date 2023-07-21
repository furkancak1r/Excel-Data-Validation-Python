# Excel Veri Doğrulama

Bu proje, Python kullanarak Excel dosyasında bulunan P1 hücresine veri doğrulama eklemek için kullanılır. Kullanıcının P1 hücresine klavyeyle veri girişi yapmasını engeller ve sadece belirli bir listeden seçim yapmasına izin verir.

## Gereksinimler

- Python 3.x
- `win32com.client` kütüphanesi
- Microsoft Excel

## Nasıl Kullanılır

1. Bu projeyi bilgisayarınıza klonlayın veya indirin.
2. "CZ-PT-A0000 Maliyetler.xlsm" adında bir Excel dosyası oluşturun veya kullanmak istediğiniz dosyayı sağlayın.
3. Python yorumlayıcısı ve `win32com.client` kütüphanesini yükleyin.
4. Script'i çalıştırarak veri doğrulama işlemini Excel dosyasındaki P1 hücresine uygulayın.
5. Kullanıcılar artık P1 hücresine klavyeyle veri girişi yapamayacak ve sadece belirtilen listeden seçebilecektir.

## Dikkat

- Script, belirtilen Excel dosyasında çalıştırılmalıdır (`file_name` değişkenini doğru dosya adıyla güncelleyin).
- `win32com.client` kütüphanesinin yüklü olduğundan emin olun. Eğer yüklü değilse, komut istemine `pip install pywin32` yazarak yükleyebilirsiniz.
- Script, veri doğrulama eklemesi yaptıktan sonra Excel dosyasını otomatik olarak kaydedecektir. Dolayısıyla, dosyayı yedekleyerek veya test verileriyle çalışarak risksiz bir şekilde deneyebilirsiniz.


# Excel Data Validation

This project is used to add data validation to cell P1 in Excel file using Python. It prevents the user from entering data into cell P1 with the keyboard and only allows selection from a specific list.

## Requirements

- Python 3.x
- `win32com.client` library
- Microsoft Excel

## How to use

1. Clone or download this project to your computer.
2. Create an Excel file named "CZ-PT-A0000 Costs.xlsm" or provide the file you want to use.
3. Install the Python interpreter and the `win32com.client` library.
4. Run the script and apply the data validation to cell P1 in the Excel file.
5. Users will no longer be able to enter data into cell P1 with the keyboard and will only be able to select from the specified list.

## Attention

- The script must be run in the specified Excel file (update the `file_name` variable with the correct filename).
- Make sure the `win32com.client` library is installed. If it is not installed, you can install it by typing `pip install pywin32` at the command prompt.
- Script will automatically save Excel file after adding data validation. So you can try it risk-free by backing up the file or working with test data.

