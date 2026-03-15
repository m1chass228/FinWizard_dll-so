//calc_salary_birulevo

#ifndef CALC_SALARY_BIRULEVO_H
#define CALC_SALARY_BIRULEVO_H

#include "iconfig.h"
#include "xlsxdocument.h"       // из QXlsx/header
#include "xlsxworksheet.h"      // из QXlsx/header
#include "xlsxformat.h"         // для форматирования

#include <QDebug>
#include <QDir>
#include <QFileInfo>
#include <QDateTime>
#include <QString>
#include <QList>

class CalcBirulevo : public QObject, public IConfig
{
    Q_OBJECT
    Q_INTERFACES(IConfig)
    Q_PLUGIN_METADATA(IID "org.finwizard.IConfig" FILE "manifest.json")

public:
    QString name() const override { return "Читает xlsx файлы и создает на их основе выжимку"; }
    QVariantMap execute(const QVariantMap &params) override;
};

#endif // CALC_SALARY_BIRULEVO_H
