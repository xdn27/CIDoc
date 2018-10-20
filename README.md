# CIDoc
CodeIgniter libary helper for exporting data to multiple document format

Library ini digunakan untuk membantu export data dokumen dengan output HTML, spreadsheet/excel atau PDF.
Membutuhkan dependency [PhpOffice\PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) untuk mengolah shpresheet dan [Dompdf](https://github.com/dompdf/dompdf) agar bisa dibuat PDF.

Saya hanya gunakan pada CodeIgniter versi 3 dan pada PHP versi 5.6, seharusnya tidak masalah pada CodeIgniter lawas, selama versi PHPnya mengikuti requirement dari dependecies.

## Instalasi
Letakan file pada folder Libraries. Pastikan dependencies sudah diinstall pada foder vendor/, silahkan sesuaikan folder vendor pada file libary bagian line require.

```
composer require phpoffice/phpspreadsheet
composer require dompdf/dompdf
```

## Contoh Penggunaan
Contoh penggunaan bisa dilihat di file controller [Example.php](https://github.com/xdn27/CIDoc/blob/master/Example.php)

> Happy coding
