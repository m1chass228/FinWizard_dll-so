// sum_columns.cpp

#include "sum_columns.h"

QVariantMap SumColumns::execute(const QVariantMap &params)
{
    QVariantMap result;

    // 1. Получаем входные параметры
    QStringList files = params.value("xlsxFiles").toStringList();
    QString outputFolder = params.value("outputFolder").toString();

    if (files.isEmpty()) {
        result["success"] = false;
        result["error"] = "Не переданы XLSX-файлы";
        return result;
    }

    if (outputFolder.isEmpty() || !QDir(outputFolder).exists()) {
        result["success"] = false;
        result["error"] = "Не указана или не существует папка для результатов";
        return result;
    }

    qDebug() << "Обработка файлов:" << files;
    qDebug() << "Папка вывода:" << outputFolder;

    double totalSum = 0.0;
    int totalRows = 0;
    QStringList processedFiles;

    // 2. Создаём новый документ для результата
    QXlsx::Document outputDoc;

    // Заголовок
    QXlsx::Format headerFormat;
    headerFormat.setFontBold(true);
    headerFormat.setFontSize(14);
    headerFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    outputDoc.write("A1", "Сумма по первому столбцу", headerFormat);
    outputDoc.mergeCells("A1:D1");

    outputDoc.write("A3", "Файл");
    outputDoc.write("B3", "Сумма столбца A");
    outputDoc.write("C3", "Количество строк");

    QXlsx::Format tableHeader;
    tableHeader.setFontBold(true);
    tableHeader.setPatternBackgroundColor(QColor(220, 220, 220));
    outputDoc.setColumnWidth(1, 40);  // ширина столбца A
    outputDoc.setColumnWidth(2, 15);
    outputDoc.setColumnWidth(3, 15);

    int row = 4;  // начиная с 4-й строки

    // 3. Обрабатываем каждый входной файл
    for (const QString &filePath : files) {
        QXlsx::Document xlsx(filePath);
        if (!xlsx.load()) {
            qWarning() << "Не удалось открыть:" << filePath;
            continue;
        }

        QXlsx::Worksheet *sheet = xlsx.currentWorksheet();
        if (!sheet) continue;

        double fileSum = 0.0;
        int fileRows = 0;

        // Читаем все строки первого столбца (A)
        for (int r = 1; r <= xlsx.dimension().rowCount(); ++r) {
            QVariant val = xlsx.read(r, 1);  // столбец A = 1
            if (val.canConvert<double>()) {
                fileSum += val.toDouble();
                fileRows++;
            }
        }

        totalSum += fileSum;
        totalRows += fileRows;
        processedFiles << filePath;

        // Записываем в результат
        outputDoc.write(row, 1, QFileInfo(filePath).fileName());
        outputDoc.write(row, 2, fileSum);
        outputDoc.write(row, 3, fileRows);
        row++;
    }

    // Итоговая строка
    QXlsx::Format totalFormat;
    totalFormat.setFontBold(true);
    totalFormat.setPatternBackgroundColor(QColor(200, 255, 200));
    outputDoc.write(row, 1, "ИТОГО", totalFormat);
    outputDoc.write(row, 2, totalSum, totalFormat);
    outputDoc.write(row, 3, totalRows, totalFormat);

    // Сохраняем
    QString timestamp = QDateTime::currentDateTime().toString("yyyy-MM-dd_hh-mm-ss");
    QString outputFile = outputFolder + "/sum_result_" + timestamp + ".xlsx";

    if (outputDoc.saveAs(outputFile)) {
        result["success"] = true;
        result["message"] = QString("Обработано файлов: %1\nОбщая сумма: %2\nСохранено в:\n%3")
                                .arg(files.size()).arg(totalSum).arg(outputFile);
        result["outputPath"] = outputFile;
    } else {
        result["success"] = false;
        result["error"] = "Не удалось сохранить результат в " + outputFile;
    }

    return result;
}

