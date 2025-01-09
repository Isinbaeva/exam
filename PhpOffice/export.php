<?php
require_once __DIR__ . '/bootstrap.php';
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

$vendorDirPath = __DIR__ . '/vendor';  
$dompdfPath = $vendorDirPath . '/dompdf/dompdf';


if (file_exists($dompdfPath)) {
    define('DOMPDF_ENABLE_AUTOLOAD', false);
    Settings::setPdfRenderer(Settings::PDF_RENDERER_DOMPDF, $vendorDirPath . '/dompdf/dompdf');
}



$publications = [
    [
        "name" => "Война и мир",
        "type" => "Книга",
        "details" => "Издательство 'Ковролин'",
        "volume" => "150 стр. / 100 стр.",
        "authors" => "Смирнова А.Г."
    ],
    [
        "name" => "Алиса в стране чудес",
        "type" => "Учебник",
        "details" => "Издательство 'Наука', 2024",
        "volume" => "250 стр. / 150 стр.",
        "authors" => "Иванов И.И., Петров П.П."
    ],
    [
        "name" => "Письма незнакомке",
        "type" => "Лабораторный практикум",
        "details" => "Издательство 'Техник', 2023",
        "volume" => "300 стр. / 200 стр.",
        "authors" => "Сидоров А.А."
    ],
    [
        "name" => "Горе от ума",
        "type" => "монография",
        "details" => "Издательство 'МГУ', 2025",
        "volume" => "200 стр. / 100 стр.",
        "authors" => "Иванова Е.Е."
    ],
    [
        "name" => "1984",
        "type" => "учебное пособие",
        "details" => "Издательство 'Прогресс', 2024",
        "volume" => "350 стр. / 250 стр.",
        "authors" => "Кузнецов В.В."
    ]
];


    // Создание документа Word
  	$phpWord = new \PhpOffice\PhpWord\PhpWord();
  	$phpWord->addParagraphStyle('p2Style', array('align'=>'center', 'spaceAfter'=>0));
  	$phpWord->addParagraphStyle('inTableStyle', array('align' => 'left', 'size' => 11, 'spaceAfter' => 0, 'spaceBefore' => 0, 'lineHeight' => 1.0, 'spacing' => 0 ));  
  
         // Добавляем текст сверху по центру
        $section = $phpWord->addSection(); // Создаем новый раздел для текста

        // Добавляем текст в центр страницы
        $section->addText(
            'СПИСОК', 
            array('align' => 'center', 'bold' => true, 'size' => 14), // выравнивание по центру, полужирный шрифт, размер шрифта
            'p2Style'
        );
        // Добавляем текст в центр страницы
        $section->addText(
            'опубликованных учебных изданий и научных трудов', 
            array('align' => 'center', 'bold' => true, 'size' => 14), // выравнивание по центру, полужирный шрифт, размер шрифта
            'p2Style'
        );
        $section->addText(
            "Вильданова А.Н.", 
            array('align' => 'center', 'bold' => true, 'size' => 12), // выравнивание по центру, полужирный шрифт, размер шрифта
            'p2Style'
        );

        // Добавляем пустую строку или другие элементы ниже
        $section->addTextBreak(1); // Перенос строки

  
 // Получаем нижний колонтитул
$footer = $section->addFooter();

$textrun = $footer->addTextRun(); 
// Добавляем текст в нижний колонтитул

$textrun = $footer->addTextRun(); 
// Добавляем текст в нижний колонтитул
$textrun->addText(
    "Автор документа Вильданов А.Н. ",
    array('align' => 'left', 'size' => 11, 'spaceAfter' => 0, 'spaceBefore' => 0, 'bold' => false)
);
  

  
  //  $section = $phpWord->addSection();
  
    // Define table style with borders
    $tableStyle = [
        'borderSize' => 6,   // Thickness of the border
        'borderColor' => '000000', // Black color for the border
        'cellMargin' => 80,  // Add some space between cell content and border
    ];

    // Define cell style with borders
    $cellStyle = [
        'borderSize' => 6,   // Border thickness for cells
        'borderColor' => '000000', // Black color for cell border
    ];

     // Apply the table style
    $table = $section->addTable($tableStyle);
  
    $table->addRow();  // Добавляем новую строку

      // Массив для ширины столбцов
      $columnWidths = [
          250,   // Ширина для первого столбца (номер)
          750,   // Ширина для второго столбца (название)
          1750,  // Ширина для третьего столбца (формат работы)
          2250,  // Ширина для четвертого столбца (выходные данные)
          1050,  // Ширина для пятого столбца (объем)
          1050   // Ширина для шестого столбца (соавторы)
      ];  
  
    // Добавляем ячейки в строку  

    $table->addCell($columnWidths[0])->addText('№ п/п', array('bold' => true )); // Примерная ширина ячейки
    $table->addCell($columnWidths[1])->addText('Название книги', array('bold' => true ));
    $table->addCell($columnWidths[2])->addText('Автор', array('bold' => true ));
    $table->addCell($columnWidths[3])->addText('Год издания', array('bold' => true ));
    $table->addCell($columnWidths[4])->addText('Объем общий в стр. или п.л. / объем, принадлежащий соискателю', array('bold' => true ));
    $table->addCell($columnWidths[5])->addText('Страна', array('bold' => true ));


    $numer = 1;

    foreach ($publications as $publication) {
        $table->addRow();
        $table->addCell($columnWidths[0], $cellStyle)->addText(strval($numer) . ".", array(), 'inTableStyle');
        $table->addCell($columnWidths[1], $cellStyle)->addText($publication['name'], array(), 'inTableStyle');
        $table->addCell($columnWidths[2], $cellStyle)->addText($publication['type'], array(), 'inTableStyle');
        $table->addCell($columnWidths[3], $cellStyle)->addText($publication['details'], array(), 'inTableStyle');
        $table->addCell($columnWidths[4], $cellStyle)->addText($publication['volume'], array(), 'inTableStyle');
        $table->addCell($columnWidths[5], $cellStyle)->addText($publication['authors'], array(), 'inTableStyle');
        $numer++;
    }

    // Сохранение документа Word
    $fileName = "Отчет.docx";
    $filePath = "results/$fileName"; // Путь, где сохранить файл
    $phpWord->save($filePath, 'Word2007');


    echo "<h2>Отчет успешно сформирован!</h2>";
        echo "<br>";
    echo "<a href='$filePath' download>Скачать отчет</a>";


?>