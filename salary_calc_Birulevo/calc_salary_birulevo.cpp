// calc_birulevo.cpp

#include "calc_salary_birulevo.h"

QVariantMap CalcBirulevo::execute(const QVariantMap &params)
{
    QVariantMap result;

    // 1. Получаем входные параметры
    QStringList files = params.value("xlsxFiles").toStringList();
    QString outputFolder = params.value("outputFolder").toString();
    QString outputFile = params.value("outputFile").toString();

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
    qDebug() << "Выходной файл:" << outputFile;

    // 2. Создаём новый документ для результата
    QXlsx::Document outputDoc;

    // Список фамилий (уже созданных листов)
    QList<QString> surnames;

    // Для каждой фамилии храним номер следующей строки для записи данных
    QMap<QString, int> surnameNextRow;

    // ============================================================
    // Определяем стили ОДИН РАЗ, до цикла
    // ============================================================

    // --- Стиль заголовка (фамилия): голубой фон, жирный, по центру ---
    QXlsx::Format headerFormat;
    headerFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    headerFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    headerFormat.setFontBold(true);
    headerFormat.setFontSize(14);
    headerFormat.setPatternBackgroundColor(QColor(173, 216, 230)); // Голубой
    headerFormat.setFontColor(QColor(0, 0, 0));
    // Рамка для заголовка
    headerFormat.setBottomBorderStyle(QXlsx::Format::BorderThin);
    headerFormat.setTopBorderStyle(QXlsx::Format::BorderThin);
    headerFormat.setLeftBorderStyle(QXlsx::Format::BorderThin);
    headerFormat.setRightBorderStyle(QXlsx::Format::BorderThin);

    // --- Стиль шапки столбцов: жирный, по центру ---
    QXlsx::Format colHeaderFormat;
    colHeaderFormat.setFontBold(true);
    colHeaderFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    colHeaderFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    colHeaderFormat.setBottomBorderStyle(QXlsx::Format::BorderThin);
    colHeaderFormat.setTopBorderStyle(QXlsx::Format::BorderThin);
    colHeaderFormat.setLeftBorderStyle(QXlsx::Format::BorderThin);
    colHeaderFormat.setRightBorderStyle(QXlsx::Format::BorderThin);

    // --- Стиль для обычных данных ---
    QXlsx::Format dataFormat;
    dataFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    dataFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    // --- Стиль для сумм: жёлтый фон, чёрная рамка, жирный ---
    QXlsx::Format sumFormat;
    sumFormat.setFontBold(true);
    sumFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    sumFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    sumFormat.setPatternBackgroundColor(QColor(255, 255, 0)); // Жёлтый
    sumFormat.setFontColor(QColor(0, 0, 0));
    sumFormat.setBottomBorderStyle(QXlsx::Format::BorderMedium);
    sumFormat.setTopBorderStyle(QXlsx::Format::BorderMedium);
    sumFormat.setLeftBorderStyle(QXlsx::Format::BorderMedium);
    sumFormat.setRightBorderStyle(QXlsx::Format::BorderMedium);

    // --- Стиль для пустых голубых ячеек в строке заголовка ---
    QXlsx::Format headerFillFormat;
    headerFillFormat.setPatternBackgroundColor(QColor(173, 216, 230));
    headerFillFormat.setBottomBorderStyle(QXlsx::Format::BorderThin);
    headerFillFormat.setTopBorderStyle(QXlsx::Format::BorderThin);
    headerFillFormat.setLeftBorderStyle(QXlsx::Format::BorderThin);
    headerFillFormat.setRightBorderStyle(QXlsx::Format::BorderThin);

    // ============================================================
    // 3. Открываем файлы
    // ============================================================
    for (const QString &filepath : files) {
        QXlsx::Document xlsx(filepath);
        if (!xlsx.load()) {
            qWarning() << "Не удалось открыть:" << filepath;
            continue;
        }

        // ======== Достаём дату из E1 ========
        double dateVal = 0.0;
        QVariant dateVar = xlsx.read(1, 5);
        if (dateVar.canConvert<double>()) {
            dateVal = dateVar.toDouble();
        }

        // Достаём месяц: дробная часть * 100
        int monthNum = qRound((dateVal - (int)dateVal) * 100);

        // Достаём год из I1
        QString year = "";
        QVariant yearVar = xlsx.read(1, 9); // I1
        if (yearVar.canConvert<QString>()) {
            year = yearVar.toString();
            year.chop(2); // убираем "г."
        }

        // Название месяца на русском
        QLocale russian(QLocale::Russian);
        QString monthName = russian.monthName(monthNum);
        if (!monthName.isEmpty()) {
            monthName[0] = monthName[0].toUpper();
        }

        // ======== Читаем фамилии из H4:H14 (столбец 8, строки 4-14) ========
        for (int i = 4; i <= 14; i++) {
            QVariant surnameVar = xlsx.read(i, 8); // столбец H
            if (!surnameVar.isValid() || !surnameVar.canConvert<QString>())
                continue;

            QString surname = surnameVar.toString().trimmed();
            if (surname.isEmpty())
                continue;

            // === Если фамилия новая — создаём лист и шапку ===
            if (!surnames.contains(surname)) {
                outputDoc.addSheet(surname);
                surnames.append(surname);
                surnameNextRow[surname] = 3; // данные начинаются с 3-й строки

                outputDoc.selectSheet(surname);

                // Объединяем B1:F1 и пишем фамилию (голубой фон)
                outputDoc.mergeCells("B1:F1");
                outputDoc.write("B1", surname, headerFormat);

                // A1 тоже голубой (чтобы строка выглядела цельно)
                outputDoc.write("A1", "", headerFillFormat);
                // G1 тоже голубой
                outputDoc.write("G1", "", headerFillFormat);

                // Шапка столбцов (строка 2) — жирным
                outputDoc.write(2, 1, "Дата",       colHeaderFormat); // A2
                outputDoc.write(2, 2, "З/п Общая",  colHeaderFormat); // B2
                outputDoc.write(2, 3, "УС",         colHeaderFormat); // C2
                outputDoc.write(2, 4, "Аптека",     colHeaderFormat); // D2
                outputDoc.write(2, 5, "Вычет",      colHeaderFormat); // E2
                outputDoc.write(2, 6, "Время",      colHeaderFormat); // F2
                outputDoc.write(2, 7, "Итого",      colHeaderFormat); // G2

                // Устанавливаем ширину столбцов
                outputDoc.setColumnWidth(1, 18.0); // Дата
                outputDoc.setColumnWidth(2, 14.0); // З/п Общая
                outputDoc.setColumnWidth(3, 14.0); // УС
                outputDoc.setColumnWidth(4, 14.0); // Аптека
                outputDoc.setColumnWidth(5, 14.0); // Вычет
                outputDoc.setColumnWidth(6, 14.0); // Время
                outputDoc.setColumnWidth(7, 14.0); // Итого
            }

            // === Переключаемся на лист этой фамилии ===
            outputDoc.selectSheet(surname);
            int row = surnameNextRow[surname];

            // ------ A: Дата (месяц + год) ------
            QString dateStr = monthName + " " + year;
            outputDoc.write(row, 1, dateStr, dataFormat);

            // ------ B: З/п Общая ------
            double zpObshaya = 0.0;
            QVariant zpVar = xlsx.read(i, 6); // столбец F той же строки
            if (zpVar.canConvert<double>()) {
                zpObshaya = zpVar.toDouble();
            }
            outputDoc.write(row, 2, zpObshaya, dataFormat);

            // ------ C: УС ------
            double usVal = 0.0;
            for (int r = 31; r <= 35; r++) {
                QVariant nameVar = xlsx.read(r, 1); // столбец A
                if (nameVar.canConvert<QString>()) {
                    QString name = nameVar.toString().trimmed();
                    if (name == surname) {
                        QVariant dVar = xlsx.read(r, 4); // столбец D
                        if (dVar.canConvert<double>()) {
                            usVal = dVar.toDouble();
                        }
                        break;
                    }
                }
            }
            outputDoc.write(row, 3, usVal, dataFormat);

            // ------ D: Аптека ------
            double aptekaVal = 0.0;
            for (int r = 40; r <= 45; r++) {
                QVariant nameVar = xlsx.read(r, 1); // столбец A
                if (nameVar.canConvert<QString>()) {
                    QString name = nameVar.toString().trimmed();
                    if (name == surname) {
                        QVariant bVar = xlsx.read(r, 2); // столбец B
                        if (bVar.canConvert<double>()) {
                            aptekaVal = bVar.toDouble();
                        }
                        break;
                    }
                }
            }
            outputDoc.write(row, 4, aptekaVal, dataFormat);

            // ------ E: Вычет (пусто) ------
            outputDoc.write(row, 5, "", dataFormat);

            // ------ F: Время ------
            QString vremya = "";
            QVariant vremyaVar = xlsx.read(i, 9); // столбец I
            if (vremyaVar.canConvert<QString>()) {
                vremya = vremyaVar.toString().trimmed();
            } else if (vremyaVar.canConvert<double>()) {
                vremya = QString::number(vremyaVar.toDouble());
            }
            outputDoc.write(row, 6, vremya, dataFormat);

            // ------ G: Итого = З/п Общая + УС ------
            double itogo = zpObshaya + usVal;
            outputDoc.write(row, 7, itogo, dataFormat);

            // Увеличиваем номер строки для этой фамилии
            surnameNextRow[surname] = row + 1;
        }
    }

    // ============================================================
    // 4. Добавляем строку с суммами на каждом листе
    // ============================================================
    for (const QString &surname : surnames) {
        outputDoc.selectSheet(surname);

        int sumRow = surnameNextRow[surname]; // следующая свободная строка

        // Подпись
        outputDoc.write(sumRow, 1, "ИТОГО:", sumFormat);

        // Считаем суммы по столбцам B, C, D, G (строки 3 .. sumRow-1)
        double totalZp = 0.0;
        double totalUs = 0.0;
        double totalApteka = 0.0;
        double totalItogo = 0.0;

        for (int r = 3; r < sumRow; r++) {
            QVariant valB = outputDoc.read(r, 2);
            if (valB.canConvert<double>()) totalZp += valB.toDouble();

            QVariant valC = outputDoc.read(r, 3);
            if (valC.canConvert<double>()) totalUs += valC.toDouble();

            QVariant valD = outputDoc.read(r, 4);
            if (valD.canConvert<double>()) totalApteka += valD.toDouble();

            QVariant valG = outputDoc.read(r, 7);
            if (valG.canConvert<double>()) totalItogo += valG.toDouble();
        }

        // Записываем суммы (жёлтый фон, чёрная рамка)
        outputDoc.write(sumRow, 2, totalZp,     sumFormat); // З/п Общая
        outputDoc.write(sumRow, 3, totalUs,      sumFormat); // УС
        outputDoc.write(sumRow, 4, totalApteka,  sumFormat); // Аптека
        outputDoc.write(sumRow, 5, "",           sumFormat); // Вычет (пусто, но стиль)
        outputDoc.write(sumRow, 6, "",           sumFormat); // Время (пусто, но стиль)
        outputDoc.write(sumRow, 7, totalItogo,   sumFormat); // Итого
    }

    // ============================================================
    // 5. Сохраняем результат
    // ============================================================
    int totalProcessed = files.size();

    if (outputDoc.saveAs(outputFile)) {
        result["success"] = true;
        result["message"] = QString("Обработано файлов: %1\nСоздано листов (фамилий): %2\nСохранено в:\n%3")
                                .arg(totalProcessed)
                                .arg(surnames.size())
                                .arg(outputFile);
        result["outputPath"] = outputFile;
    } else {
        result["success"] = false;
        result["error"] = "Не удалось сохранить результат в " + outputFile;
    }

    return result;
}
