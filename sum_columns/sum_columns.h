// sum_columns.h

#ifndef SUM_COLUMNS_H
#define SUM_COLUMNS_H

#include "iconfig.h"
#include "xlsxdocument.h"       // из QXlsx/header
#include "xlsxworksheet.h"      // из QXlsx/header
#include "xlsxformat.h"         // для форматирования

#include <QDebug>
#include <QDir>
#include <QFileInfo>
#include <QDateTime>

class SumColumns : public QObject, public IConfig
{
    Q_OBJECT
    Q_INTERFACES(IConfig)
    Q_PLUGIN_METADATA(IID "org.finwizard.IConfig" FILE "manifest.json")

public:
    QString name() const override { return "Сумма столбцов XLSX"; }
    QVariantMap execute(const QVariantMap &params) override;
};

#endif // SUM_COLUMNS_H
