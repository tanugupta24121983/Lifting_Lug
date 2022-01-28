#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QDebug>
#include <QFileDialog>

#include <QMap>
#include <QtMath>
#include <QCompleter>
#include <QSettings>
#include <QMessageBox>


QMap<QString, QMap<QString, QMap<QString, QList<QString>>>> m_map;
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
    ,excel(nullptr)
    ,workbooks(nullptr)
    ,workbook(nullptr)
    ,sheets(nullptr)
    ,sheet(nullptr)
{
    ui->setupUi(this);
    QPixmap pixmap(":/images/images/Isgec-logo.jpg");
    int w = ui->label_pic->width();
    int h = ui->label_pic->height();
    // DO FOR ALL WINDOWS
    ui->label_pic->setPixmap(pixmap.scaled(w,h,Qt::KeepAspectRatio));
    ui->label_pic_2->setPixmap(pixmap.scaled(w,h,Qt::KeepAspectRatio));
    ui->label_pic_5->setPixmap(pixmap.scaled(w,h,Qt::KeepAspectRatio));
     // ENDS HERE
    QPixmap pixmap1(":/images/images/LIFTING.jpg");
    w = ui->label_pic2->width();
    h = ui->label_pic2->height();

    ui->label_pic2->setPixmap(pixmap1.scaled(w,h,Qt::KeepAspectRatio));



    QPixmap pixmap2(":/images/images/FLOATING_HEAD.jpg");
    w = ui->label_pic2_2->width();
    h = ui->label_pic2_2->height();

    ui->label_pic2_2->setPixmap(pixmap2.scaled(w,h,Qt::KeepAspectRatio));

    QPixmap pixmap4(":/images/images/Lifting_lug_channel_cover.jpg");
    w = ui->label_pic2_5->width();
    h = ui->label_pic2_5->height();

    ui->label_pic2_5->setPixmap(pixmap4.scaled(w,h,Qt::KeepAspectRatio));


    ui->le_temprature->setText(QString::number(100, 'g', 3));
    ui->cb_lifting_ug_materia->setInsertPolicy(QComboBox::NoInsert);
    ui->cb_lifting_ug_materia->completer()->setCompletionMode(QCompleter::PopupCompletion);
    ui->cb_lifting_ug_materia->completer()->setFilterMode(Qt::MatchContains);

    ui->fh_cb_lifting_ug_materia->setInsertPolicy(QComboBox::NoInsert);
    ui->fh_cb_lifting_ug_materia->completer()->setCompletionMode(QCompleter::PopupCompletion);
    ui->fh_cb_lifting_ug_materia->completer()->setFilterMode(Qt::MatchContains);

    ui->fh_cb_lifting_ug_materia_2->setInsertPolicy(QComboBox::NoInsert);
    ui->fh_cb_lifting_ug_materia_2->completer()->setCompletionMode(QCompleter::PopupCompletion);
    ui->fh_cb_lifting_ug_materia_2->completer()->setFilterMode(Qt::MatchContains);

    ui->le_temprature->setText("100");
    load_temprature_data();
    ui->stackedWidget->setCurrentIndex(0);

    connect(ui->cb_lifting_ug_materia,SIGNAL(textActivated(QString)), this, SLOT (on_cb_lifting_ug_materia_activated(QString)));
    connect(ui->cb_grade,SIGNAL(textActivated(QString)), this, SLOT (on_cb_grade_activated(QString)));
    connect(ui->cb_thickness,SIGNAL(textActivated(QString)), this, SLOT (on_cb_thickness_activated(QString)));

    connect(ui->fh_cb_lifting_ug_materia,SIGNAL(textActivated(QString)), this, SLOT (on_fh_cb_lifting_ug_materia_activated(QString)));
    connect(ui->fh_cb_grade,SIGNAL(textActivated(QString)), this, SLOT (on_fh_cb_grade_activated(QString)));
    connect(ui->fh_cb_thicknes,SIGNAL(textActivated(QString)), this, SLOT (on_fh_cb_thicknes_activated(QString)));
}

void MainWindow::saveSettings() {
    QSettings settings("settings.ini", QSettings::IniFormat);
    settings.setValue("le_wt_to_be_lifted", ui->le_wt_to_be_lifted->text());
    settings.setValue("le_shock_factor", ui->le_shock_factor->text());
    settings.setValue("le_no_of_liftin_lugs", ui->le_no_of_liftin_lugs->text());
}

void MainWindow::loadSettings() {
    QSettings settings("settings.ini", QSettings::IniFormat);
    ui->le_wt_to_be_lifted->setText( settings.value("le_wt_to_be_lifted").toString());
    ui->le_shock_factor->setText( settings.value("le_shock_factor").toString());
    ui->le_no_of_liftin_lugs->setText( settings.value("le_no_of_liftin_lugs").toString());
}
MainWindow::~MainWindow()
{
    // clean up and close up
    //saveSettings();
    delete ui;
}

void MainWindow::load_temprature_data()
{
    QString fileName = "C:/ISGEC-TOOLS/yield stress excel.xlsx";
    QFile file(fileName);
    try {
        if(file.open(QIODevice::ReadOnly)) {
            qDebug() << "File opened succesfully";
            auto excel     = new QAxObject("Excel.Application");
            auto workbooks = excel->querySubObject("Workbooks");
            auto workbook  = workbooks->querySubObject("Open(const QString&)",fileName);
            auto sheets    = workbook->querySubObject("Worksheets");
            auto sheet     = sheets->querySubObject("Item(int)", 1);

            QVariant var;
            if (sheet != NULL && ! sheet->isNull())
            {
                QAxObject *usedRange = sheet->querySubObject("UsedRange");

                var = usedRange->dynamicCall("Value");
                delete usedRange;
            }

            workbook->dynamicCall("Close()");
            excel->dynamicCall("Quit()");
            delete excel;

            QVariantList varRows = var.toList();
            const int rowCount = varRows.size();
            QVariantList rowData;

            for(int i=1;i<rowCount;++i)
            {
                rowData = varRows[i].toList();
                auto specno =rowData[2].toString();
                specno = specno.trimmed();
                if(!m_map.contains(specno)) {
                    auto grade= rowData[3].toString();
                    m_map.insert(specno,{});
                    QMap<QString, QMap<QString, QList<QString>>> &it = m_map.find(specno).value();
                    it.insert(grade,{});
                    QMap<QString, QList<QString>> &it_grade = it.find(grade).value();
                    auto thickness = rowData[4].toString();
                    QList<QString> temp_list;
                    auto composition = rowData[1].toString();
                    temp_list.append(composition);
                    for(int i = 0; i <= 20; i++)
                    {
                        auto temprature = rowData[5 + i].toString();
                        temp_list.append(temprature);
                    }

                    it_grade.insert(thickness,temp_list);
                }

                else {
                    QMap<QString, QMap<QString, QList<QString>>> &it = m_map.find(specno).value();
                    auto grade = rowData[3].toString();
                    grade = grade.trimmed();
                    QList<QString> temp_list;
                    auto composition = rowData[1].toString();
                    temp_list.append(composition);
                    for(int i = 0; i <= 20; i++)
                    {
                        auto temprature = rowData[5 + i].toString();
                        temp_list.append(temprature);
                    }
                    if(!it.contains(grade)) {
                        it.insert(grade,{});
                        QMap<QString, QList<QString>> &it_grade = it.find(grade).value();
                        auto thickness = rowData[4].toString();
                        it_grade.insert(thickness,temp_list);
                    }
                    else {
                        QMap<QString, QList<QString>> &it_grade = it.find(grade).value();
                        auto thickness = rowData[4].toString();
                        it_grade.insert(thickness,temp_list);
                    }
                }
            }

        }
    }
    catch(...) {
        file.close();
    }
    file.close();
    QMapIterator<QString,QMap<QString, QMap<QString, QList<QString>>>> i(m_map);
    while (i.hasNext()) {
        i.next();
        ui->cb_lifting_ug_materia->addItem(i.key());
        ui->fh_cb_lifting_ug_materia->addItem(i.key());   
        ui->fh_cb_lifting_ug_materia_2->addItem(i.key());
    }
    ui->cb_lifting_ug_materia->setCurrentIndex(0);
    ui->fh_cb_lifting_ug_materia->setCurrentIndex(0);
    ui->fh_cb_lifting_ug_materia_2->setCurrentIndex(0);
}
void MainWindow::on_pb_check_for_shear_clicked()
{
    double shear_stress_of_Log = 0.4 * (ui->le_sy->text().toDouble());
    ui->le_shear_stree_of_Log->setText(QString::number(shear_stress_of_Log, 'g', 3));
    double h32 = ui->le_max_wt_per_lifting_lugs->text().toDouble();
    double h39 = ui->le_distance_of_lifting_lug_hole_to_top->text().toDouble();
    double h43 = ui->le_diameter_of_hole->text().toDouble();
    double quotient = (((h32 *9.81)/1000)*1000*1000);
    double reqd_thickness= quotient /((2*(shear_stress_of_Log*100)*98.0662)*(h39 - h43/2));
    ui->le_reqd_thickness->setText(QString::number(reqd_thickness, 'g', 3));
    ui->le_provided_thickness->setText(ui->le_thickness_of_lug->text());
    if(ui->le_provided_thickness->text().toDouble() > reqd_thickness ) {
        ui->cb_safe->setChecked(true);
        ui->cb_safe->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->cb_safe->setChecked(false);
        ui->cb_safe->setStyleSheet("background-color: rgb(255, 0, 0);");
    }
}

void MainWindow::on_pb_check_for_bending_clicked()
{
    double bendnig_stress_of_Log = 0.667 * (ui->le_sy->text().toDouble());
    ui->le_bending_stree_of_Log->setText(QString::number(bendnig_stress_of_Log, 'g', 3));

    double h32 = ui->le_max_wt_per_lifting_lugs->text().toDouble();
    double h42 = ui->le_dis_of_centre_line_of_hole->text().toDouble();
    double h40 = ui->le_width_of_log->text().toDouble();
    double reqd_thickness =(6*((h32*9.81)/1000)*h42)*1000*1000/(((bendnig_stress_of_Log*100)*98.0662)*((h40)*(h40)));
    ui->le_reqd_thickness_for_bending->setText(QString::number(reqd_thickness, 'g', 3));
    ui->le_provided_thickness_bending->setText(ui->le_thickness_of_lug->text());
    if(ui->le_provided_thickness_bending->text().toDouble() > reqd_thickness ) {
        ui->cb_safe_bending->setChecked(true);
        ui->cb_safe_bending->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->cb_safe_bending->setChecked(false);
        ui->cb_safe_bending->setStyleSheet("background-color: rgb(255, 0, 0);");
    }
}

void MainWindow::on_le_wt_to_be_lifted_editingFinished()
{
    double wt_to_be_lifted = ui->le_wt_to_be_lifted->text().toDouble();
    double shock_factor = ui->le_shock_factor->text().toDouble();
    double no_of_liftin_lugs = ui->le_no_of_liftin_lugs->text().toDouble();

    double max_wt_per_lifting_lugs= ((wt_to_be_lifted * shock_factor ) / no_of_liftin_lugs);
    ui->le_max_wt_per_lifting_lugs->setText(QString::number(max_wt_per_lifting_lugs));
}

void MainWindow::on_le_shock_factor_editingFinished()
{
    double wt_to_be_lifted = ui->le_wt_to_be_lifted->text().toDouble();
    double shock_factor = ui->le_shock_factor->text().toDouble();
    double no_of_liftin_lugs = ui->le_no_of_liftin_lugs->text().toDouble();

    double max_wt_per_lifting_lugs= ((wt_to_be_lifted * shock_factor ) / no_of_liftin_lugs);
    ui->le_max_wt_per_lifting_lugs->setText(QString::number(max_wt_per_lifting_lugs));
}

void MainWindow::on_le_no_of_liftin_lugs_editingFinished()
{
    double wt_to_be_lifted = ui->le_wt_to_be_lifted->text().toDouble();
    double shock_factor = ui->le_shock_factor->text().toDouble();
    double no_of_liftin_lugs = ui->le_no_of_liftin_lugs->text().toDouble();

    double max_wt_per_lifting_lugs= ((wt_to_be_lifted * shock_factor ) / no_of_liftin_lugs);
    ui->le_max_wt_per_lifting_lugs->setText(QString::number(max_wt_per_lifting_lugs));
}


void MainWindow::on_cb_lifting_ug_materia_activated(const QString &arg1)
{
    QMap<QString, QMap<QString, QList<QString>>> &it_grade = m_map.find(arg1).value();
    ui->cb_grade->clear();
    if(!it_grade.empty()) {
        QMapIterator<QString, QMap<QString, QList<QString>>> i(it_grade);
        while (i.hasNext()) {
            i.next();
            QString key = i.key();
            ui->cb_grade->addItem(key);
            ui->cb_grade->setCurrentIndex(-1);
        }
    }
    ui->cb_grade->setCurrentIndex(0);
    if(ui->cb_grade->count())
        on_cb_grade_activated_custom(ui->cb_grade->currentText());
}


void MainWindow::on_fh_cb_lifting_ug_materia_activated(const QString &arg1)
{
    QMap<QString, QMap<QString, QList<QString>>> &it_grade = m_map.find(arg1).value();
    ui->fh_cb_grade->clear();
    if(!it_grade.empty()) {
        QMapIterator<QString, QMap<QString, QList<QString>>> i(it_grade);
        while (i.hasNext()) {
            i.next();
            QString key = i.key();
            ui->fh_cb_grade->addItem(key);
            ui->fh_cb_grade->setCurrentIndex(-1);
        }
    }
    ui->fh_cb_grade->setCurrentIndex(0);
    if(ui->fh_cb_grade->count())
        on_fh_cb_grade_activated_custom(ui->fh_cb_grade->currentText());
}

void MainWindow::on_fh_cb_lifting_ug_materia_2_textActivated(const QString &arg1)
{
    QMap<QString, QMap<QString, QList<QString>>> &it_grade = m_map.find(arg1).value();
    ui->fh_cb_grade_2->clear();
    if(!it_grade.empty()) {
        QMapIterator<QString, QMap<QString, QList<QString>>> i(it_grade);
        while (i.hasNext()) {
            i.next();
            QString key = i.key();
            ui->fh_cb_grade_2->addItem(key);
            ui->fh_cb_grade_2->setCurrentIndex(-1);
        }
    }
    ui->fh_cb_grade_2->setCurrentIndex(0);
    if(ui->fh_cb_grade_2->count())
        on_fh_cb_grade_2_activated_custom(ui->fh_cb_grade_2->currentText());

}

void MainWindow::on_cb_thickness_activated(const QString &arg1)
{
    QList<QString> &ref = m_map.find(ui->cb_lifting_ug_materia->currentText()).value().find(ui->cb_grade->currentText()).value().find(arg1).value();
    if(!ref.isEmpty()) {
        ui->le_temprature->setText("100");
        ui->le_composition->setText(ref.at(0));
        ui->le_uns_no->setText(ref.at(1));
        ui->le_product_type->setText(ref.at(2));
        ui->le_sy->setText(QString::number( ref.at(3).toDouble() *0.70306957964239));
    }
}

void MainWindow::on_fh_cb_thicknes_activated(const QString &arg1)
{
    QList<QString> &ref = m_map.find(ui->fh_cb_lifting_ug_materia->currentText()).value().find(ui->fh_cb_grade->currentText()).value().find(arg1).value();
    if(!ref.isEmpty()) {
        ui->fh_le_tempratur->setText("100");
        ui->fh_le_composition->setText(ref.at(0));
        ui->fh_le_uns_no->setText(ref.at(1));
        ui->fh_le_product_type->setText(ref.at(2));
        ui->fh_le_sy->setText(QString::number( ref.at(3).toDouble() *0.70306957964239));
    }

}

void MainWindow::on_fh_cb_thicknes_2_activated(const QString &arg1)
{
    QList<QString> &ref = m_map.find(ui->fh_cb_lifting_ug_materia_2->currentText()).value().find(ui->fh_cb_grade_2->currentText()).value().find(arg1).value();
    if(!ref.isEmpty()) {
        ui->fh_le_tempratur_2->setText("100");
        ui->fh_le_composition_2->setText(ref.at(0));
        ui->fh_le_uns_no_2->setText(ref.at(1));
        ui->fh_le_product_type_2->setText(ref.at(2));
        ui->fh_le_sy_2->setText(QString::number( ref.at(3).toDouble() *0.70306957964239));
    }

}

void MainWindow::on_cb_grade_activated(const QString &arg1)
{
    QMap<QString, QList<QString>> &ref = m_map.find(ui->cb_lifting_ug_materia->currentText()).value().find(arg1).value();
    QMapIterator<QString, QList<QString>> i_thickness(ref);
    ui->cb_thickness->clear();
    ui->cb_thickness->setCurrentIndex(-1);
    while (i_thickness.hasNext()) {
        i_thickness.next();
        ui->cb_thickness->addItem(i_thickness.key());
    }
    ui->cb_thickness->setCurrentIndex(0);
    if(ui->cb_thickness->count())
        on_cb_thickness_activated_custom(ui->cb_thickness->currentText());
}

void MainWindow::on_fh_cb_grade_activated(const QString &arg1)
{
    QMap<QString, QList<QString>> &ref = m_map.find(ui->fh_cb_lifting_ug_materia->currentText()).value().find(arg1).value();
    QMapIterator<QString, QList<QString>> i_thickness(ref);
    ui->fh_cb_thicknes->clear();
    ui->fh_cb_thicknes->setCurrentIndex(-1);
    while (i_thickness.hasNext()) {
        i_thickness.next();
        ui->fh_cb_thicknes->addItem(i_thickness.key());
    }
    ui->fh_cb_thicknes->setCurrentIndex(0);
    if(ui->fh_cb_thicknes->count())
        on_fh_cb_thickness_activated_custom(ui->fh_cb_thicknes->currentText());
}

void MainWindow::on_fh_cb_grade_2_textActivated(const QString &arg1)
{
    QMap<QString, QList<QString>> &ref = m_map.find(ui->fh_cb_lifting_ug_materia_2->currentText()).value().find(arg1).value();
    QMapIterator<QString, QList<QString>> i_thickness(ref);
    ui->fh_cb_thicknes_2->clear();
    ui->fh_cb_thicknes_2->setCurrentIndex(-1);
    while (i_thickness.hasNext()) {
        i_thickness.next();
        ui->fh_cb_thicknes_2->addItem(i_thickness.key());
    }
    ui->fh_cb_thicknes_2->setCurrentIndex(0);
    if(ui->fh_cb_thicknes_2->count())
        on_fh_cb_2_thickness_activated_custom(ui->fh_cb_thicknes_2->currentText());

}

void MainWindow::on_cb_thickness_activated_custom(const QString &arg1)
{
    QList<QString> &ref = m_map.find(ui->cb_lifting_ug_materia->currentText()).value().find(ui->cb_grade->currentText()).value().find(arg1).value();
    if(!ref.isEmpty()) {
        ui->le_temprature->setText("100");
        ui->le_composition->setText(ref.at(0));
        ui->le_uns_no->setText(ref.at(1));
        ui->le_product_type->setText(ref.at(2));
        ui->le_sy->setText(QString::number( ref.at(3).toDouble() *0.70306957964239));
    }
}

void MainWindow::on_fh_cb_thickness_activated_custom(const QString &arg1)
{
    QList<QString> &ref = m_map.find(ui->fh_cb_lifting_ug_materia->currentText()).value().find(ui->fh_cb_grade->currentText()).value().find(arg1).value();
    if(!ref.isEmpty()) {
        ui->fh_le_tempratur->setText("100");
        ui->fh_le_composition->setText(ref.at(0));
        ui->fh_le_uns_no->setText(ref.at(1));
        ui->fh_le_product_type->setText(ref.at(2));
        ui->fh_le_sy->setText(QString::number( ref.at(3).toDouble() *0.70306957964239));
    }
}

void MainWindow::on_fh_cb_2_thickness_activated_custom(const QString &arg1)
{
    QList<QString> &ref = m_map.find(ui->fh_cb_lifting_ug_materia_2->currentText()).value().find(ui->fh_cb_grade_2->currentText()).value().find(arg1).value();
    if(!ref.isEmpty()) {
        ui->fh_le_tempratur_2->setText("100");
        ui->fh_le_composition_2->setText(ref.at(0));
        ui->fh_le_uns_no_2->setText(ref.at(1));
        ui->fh_le_product_type_2->setText(ref.at(2));
        ui->fh_le_sy_2->setText(QString::number( ref.at(3).toDouble() *0.70306957964239));
    }
}

void MainWindow::on_cb_grade_activated_custom(const QString &arg1)
{
    QMap<QString, QList<QString>> &ref = m_map.find(ui->cb_lifting_ug_materia->currentText()).value().find(arg1).value();
    QMapIterator<QString, QList<QString>> i_thickness(ref);
    ui->cb_thickness->clear();
    ui->cb_thickness->setCurrentIndex(-1);
    while (i_thickness.hasNext()) {
        i_thickness.next();
        ui->cb_thickness->addItem(i_thickness.key());
    }
    ui->cb_thickness->setCurrentIndex(0);
    if(ui->cb_thickness->count())
        on_cb_thickness_activated_custom(ui->cb_thickness->currentText());
}

void MainWindow::on_fh_cb_grade_activated_custom(const QString &arg1)
{
    QMap<QString, QList<QString>> &ref = m_map.find(ui->fh_cb_lifting_ug_materia->currentText()).value().find(arg1).value();
    QMapIterator<QString, QList<QString>> i_thickness(ref);
    ui->fh_cb_thicknes->clear();
    ui->fh_cb_thicknes->setCurrentIndex(-1);
    while (i_thickness.hasNext()) {
        i_thickness.next();
        ui->fh_cb_thicknes->addItem(i_thickness.key());
    }
    ui->fh_cb_thicknes->setCurrentIndex(0);
    if(ui->fh_cb_thicknes->count())
        on_fh_cb_thickness_activated_custom(ui->fh_cb_thicknes->currentText());
}

void MainWindow::on_fh_cb_grade_2_activated_custom(const QString &arg1)
{
    QMap<QString, QList<QString>> &ref = m_map.find(ui->fh_cb_lifting_ug_materia_2->currentText()).value().find(arg1).value();
    QMapIterator<QString, QList<QString>> i_thickness(ref);
    ui->fh_cb_thicknes_2->clear();
    ui->fh_cb_thicknes_2->setCurrentIndex(-1);
    while (i_thickness.hasNext()) {
        i_thickness.next();
        ui->fh_cb_thicknes_2->addItem(i_thickness.key());
    }
    ui->fh_cb_thicknes_2->setCurrentIndex(0);
    if(ui->fh_cb_thicknes_2->count())
        on_fh_cb_2_thickness_activated_custom(ui->fh_cb_thicknes_2->currentText());
}

void MainWindow::on_le_temprature_editingFinished()
{
    QList<QString> &ref = m_map.find(ui->cb_lifting_ug_materia->currentText()).value().find(ui->cb_grade->currentText()).value().find(ui->cb_thickness->currentText()).value();
    if(!ref.isEmpty()) {
        int temp = ui->le_temprature->text().toInt();
        temp = temp - 100;
        if(temp < 0) {

        }
        else {
            int factor = temp/50;
            int percentyl = temp % 50;

            if(percentyl == 0) {
                ui->le_sy->setText(QString::number(ref.at(factor + 3).toDouble() *0.70306957964239));
            }
            else {
                float factor = ui->le_temprature->text().toFloat() - 100.0;

                int range = factor / 50;
                int range_plus_1 = (factor / 50) +1;

                double degree_factor = (abs(ref.at(range + 3).toDouble()*0.70306957964239 - ref.at(range_plus_1 + 3).toDouble()*0.70306957964239))/50;
                double final_temp;
                if(abs(ref.at(range + 3).toDouble() < ref.at(range_plus_1 + 3).toDouble()))
                    final_temp = ref.at(range + 3).toDouble()*0.70306957964239 + (degree_factor * (temp % 50));
                else if (abs(ref.at(range + 3).toDouble()*0.70306957964239 > ref.at(range_plus_1 + 3).toDouble()*0.70306957964239))
                    final_temp = ref.at(range + 3).toDouble()*0.70306957964239 - (degree_factor * (temp % 50));
                else
                    final_temp = ref.at(range).toDouble()*0.70306957964239;

                ui->le_sy->setText(QString::number(final_temp, 'g', 3));

            }
        }
    }
}


void MainWindow::on_fh_le_tempratur_editingFinished()
{
    QList<QString> &ref = m_map.find(ui->fh_cb_lifting_ug_materia->currentText()).value().find(ui->fh_cb_grade->currentText()).value().find(ui->fh_cb_thicknes->currentText()).value();
    if(!ref.isEmpty()) {
        int temp = ui->fh_le_tempratur->text().toInt();
        temp = temp - 100;
        if(temp < 0) {

        }
        else {
            int factor = temp/50;
            int percentyl = temp % 50;

            if(percentyl == 0) {
                ui->fh_le_sy->setText(QString::number(ref.at(factor + 3).toDouble() *0.70306957964239));
            }
            else {
                float factor = ui->fh_le_tempratur->text().toFloat() - 100.0;

                int range = factor / 50;
                int range_plus_1 = (factor / 50) +1;

                double degree_factor = (abs(ref.at(range + 3).toDouble()*0.70306957964239 - ref.at(range_plus_1 + 3).toDouble()*0.70306957964239))/50;
                double final_temp;
                if(abs(ref.at(range + 3).toDouble() < ref.at(range_plus_1 + 3).toDouble()))
                    final_temp = ref.at(range + 3).toDouble()*0.70306957964239 + (degree_factor * (temp % 50));
                else if (abs(ref.at(range + 3).toDouble()*0.70306957964239 > ref.at(range_plus_1 + 3).toDouble()*0.70306957964239))
                    final_temp = ref.at(range + 3).toDouble()*0.70306957964239 - (degree_factor * (temp % 50));
                else
                    final_temp = ref.at(range).toDouble()*0.70306957964239;

                ui->fh_le_sy->setText(QString::number(final_temp, 'g', 3));

            }
        }
    }
}

void MainWindow::on_fh_le_tempratur_2_editingFinished()
{
    QList<QString> &ref = m_map.find(ui->fh_cb_lifting_ug_materia_2->currentText()).value().find(ui->fh_cb_grade_2->currentText()).value().find(ui->fh_cb_thicknes_2->currentText()).value();
    if(!ref.isEmpty()) {
        int temp = ui->fh_le_tempratur_2->text().toInt();
        temp = temp - 100;
        if(temp < 0) {

        }
        else {
            int factor = temp/50;
            int percentyl = temp % 50;

            if(percentyl == 0) {
                ui->fh_le_sy_2->setText(QString::number(ref.at(factor + 3).toDouble() *0.70306957964239));
            }
            else {
                float factor = ui->fh_le_tempratur_2->text().toFloat() - 100.0;

                int range = factor / 50;
                int range_plus_1 = (factor / 50) +1;

                double degree_factor = (abs(ref.at(range + 3).toDouble()*0.70306957964239 - ref.at(range_plus_1 + 3).toDouble()*0.70306957964239))/50;
                double final_temp;
                if(abs(ref.at(range + 3).toDouble() < ref.at(range_plus_1 + 3).toDouble()))
                    final_temp = ref.at(range + 3).toDouble()*0.70306957964239 + (degree_factor * (temp % 50));
                else if (abs(ref.at(range + 3).toDouble()*0.70306957964239 > ref.at(range_plus_1 + 3).toDouble()*0.70306957964239))
                    final_temp = ref.at(range + 3).toDouble()*0.70306957964239 - (degree_factor * (temp % 50));
                else
                    final_temp = ref.at(range).toDouble()*0.70306957964239;

                ui->fh_le_sy_2->setText(QString::number(final_temp, 'g', 3));

            }
        }
    }

}


void MainWindow::on_pb_check_for_shear_in_weld_clicked()
{
    double h40 = ui->le_width_of_log->text().toDouble();
    double total_weld_length = 2.0 * h40;
    ui->le_total_weld_length->setText(QString::number(total_weld_length, 'g', 3));
    double proided_fillet_size = ui->le_fillet_lug_to_pad->text().toDouble();
    ui->le_provided_fillet_size->setText(QString::number(proided_fillet_size, 'g', 3));
    double weld_area = total_weld_length*proided_fillet_size*qSin(3.14/4);
    ui->le_weld_area_Aw1->setText(QString::number(weld_area));
    double h32 = ui->le_max_wt_per_lifting_lugs->text().toDouble();
    double induced_shear_stress_in_weld = h32/weld_area;
    ui->le_induce_shear_stress_in_weld->setText(QString::number(induced_shear_stress_in_weld, 'g', 3));
    double allowable_equivalet_stress_Sw = 0.4*ui->le_sy->text().toFloat();
    ui->le_allowable_equivalet_stress_Sw->setText(QString::number(allowable_equivalet_stress_Sw, 'g', 3));

    if(induced_shear_stress_in_weld < ui->le_allowable_equivalet_stress_Sw->text().toDouble() ) {
        ui->cb_safe_shear_in_weld->setChecked(true);
        ui->cb_safe_shear_in_weld->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->cb_safe_shear_in_weld->setChecked(false);
        ui->cb_safe_shear_in_weld->setStyleSheet("background-color: rgb(255, 0, 0);");
    }
}

void MainWindow::on_pb_generate_report_clicked()
{
    // clean up and close up
    if(nullptr != workbook) {
        workbook->dynamicCall("Save()");
        workbook->dynamicCall("Close()");
        workbook->dynamicCall("Quit()");
    }
    if(excel != workbook)
        excel->dynamicCall("Quit()");

    QString fileName = "C:/ISGEC-TOOLS/Lifting Lug Shell Cover.xlsx";
    QString fileName_pdf = "C:/ISGEC-TOOLS/Lifting_Lug_Shell_Cover";
    QFile file_write(fileName);
    if(file_write.open(QIODevice::ReadWrite)) {
       excel     = new QAxObject("Excel.Application");
       workbooks = excel->querySubObject("Workbooks");
       workbook  = workbooks->querySubObject("Open(const QString&)",fileName);
       sheets    = workbook->querySubObject("Worksheets");
       sheet     = sheets->querySubObject("Item(int)", 1);
    }
    else {
        QMessageBox msgBox;
        msgBox.setText("Please Make Sure Excel file is closed!!");
        msgBox.exec();
        return;;
    }
    file_write.close();

    auto cell = sheet->querySubObject("Cells(int,int)", 52,8);
    cell->setProperty("Value", ui->le_shear_stree_of_Log->text());
    cell = sheet->querySubObject("Cells(int,int)", 54,8);
    cell->setProperty("Value", ui->le_reqd_thickness->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,5);
    cell->setProperty("Value", ui->le_thickness_of_lug->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,7);
    cell->setProperty("Value", ui->le_reqd_thickness->text());


    cell = sheet->querySubObject("Cells(int,int)", 70,8);
    cell->setProperty("Value", ui->le_bending_stree_of_Log->text());
    cell = sheet->querySubObject("Cells(int,int)", 71,8);
    cell->setProperty("Value", ui->le_reqd_thickness_for_bending->text());
    cell = sheet->querySubObject("Cells(int,int)", 73,5);
    cell->setProperty("Value", ui->le_provided_thickness_bending->text());
    cell = sheet->querySubObject("Cells(int,int)", 73,7);
    cell->setProperty("Value", ui->le_reqd_thickness_for_bending->text());

    // on_pb_check_for_shear_in_weld_clicked/////////////////////////////////////////////////////////////////////

    cell = sheet->querySubObject("Cells(int,int)", 77,8);
    cell->setProperty("Value", ui->le_total_weld_length->text());
    cell = sheet->querySubObject("Cells(int,int)", 79,8);
    cell->setProperty("Value", ui->le_provided_fillet_size->text());
    cell = sheet->querySubObject("Cells(int,int)", 81,8);
    cell->setProperty("Value", ui->le_weld_area_Aw1->text());
    cell = sheet->querySubObject("Cells(int,int)", 83,8);
    cell->setProperty("Value", ui->le_induce_shear_stress_in_weld->text());
    cell = sheet->querySubObject("Cells(int,int)", 85,8);
    cell->setProperty("Value", ui->le_allowable_equivalet_stress_Sw->text());


    cell = sheet->querySubObject("Cells(int,int)", 29,8);
    cell->setProperty("Value", ui->le_wt_to_be_lifted->text());
    cell = sheet->querySubObject("Cells(int,int)", 32,8);
    cell->setProperty("Value", ui->le_max_wt_per_lifting_lugs->text());

    cell = sheet->querySubObject("Cells(int,int)", 30,8);
    cell->setProperty("Value", ui->le_shock_factor->text());
    cell = sheet->querySubObject("Cells(int,int)", 32,8);
    cell->setProperty("Value", ui->le_max_wt_per_lifting_lugs->text());

    cell = sheet->querySubObject("Cells(int,int)", 31,8);
    cell->setProperty("Value", ui->le_no_of_liftin_lugs->text());
    cell = sheet->querySubObject("Cells(int,int)", 32,8);
    cell->setProperty("Value", ui->le_max_wt_per_lifting_lugs->text());

    cell = sheet->querySubObject("Cells(int,int)", 36,8);
    cell->setProperty("Value", ui->le_sy->text());

    cell = sheet->querySubObject("Cells(int,int)", 39,8);
    cell->setProperty("Value", ui->le_distance_of_lifting_lug_hole_to_top->text());

    cell = sheet->querySubObject("Cells(int,int)", 40,8);
    cell->setProperty("Value", ui->le_width_of_log->text());

    cell = sheet->querySubObject("Cells(int,int)", 41,8);
    cell->setProperty("Value", ui->le_thickness_of_lug->text());

    cell = sheet->querySubObject("Cells(int,int)", 42,8);
    cell->setProperty("Value", ui->le_dis_of_centre_line_of_hole->text());

    cell = sheet->querySubObject("Cells(int,int)", 43,8);
    cell->setProperty("Value", ui->le_diameter_of_hole->text());

    cell = sheet->querySubObject("Cells(int,int)", 44,8);
    cell->setProperty("Value", ui->le_fillet_lug_to_pad->text());

    cell = sheet->querySubObject("Cells(int,int)", 45,8);
    cell->setProperty("Value", ui->le_reinforcing_pad_l1->text());

    cell = sheet->querySubObject("Cells(int,int)", 46,8);
    cell->setProperty("Value", ui->le_reinforcing_pad_l2->text());

    cell = sheet->querySubObject("Cells(int,int)", 47,8);
    cell->setProperty("Value", ui->le_thickness_of_pad->text());

    cell = sheet->querySubObject("Cells(int,int)", 48,8);
    cell->setProperty("Value", ui->le_fillet_weld_leg_size_f2->text());

    cell = sheet->querySubObject("Cells(int,int)", 35,8);
    cell->setProperty("Value", ui->cb_lifting_ug_materia->currentText());
    cell = sheet->querySubObject("Cells(int,int)", 35,9);
    cell->setProperty("Value", ui->cb_grade->currentText());



    cell = sheet->querySubObject("Cells(int,int)", 5,3);
    cell->setProperty("Value", ui->le_designed_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 67,3);
    cell->setProperty("Value", ui->le_designed_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 1,5);
    cell->setProperty("Value", ui->le_client->text());

    cell = sheet->querySubObject("Cells(int,int)", 63,5);
    cell->setProperty("Value", ui->le_client->text());

    cell = sheet->querySubObject("Cells(int,int)", 3,5);
    cell->setProperty("Value", ui->le_eqpt->text());

    cell = sheet->querySubObject("Cells(int,int)", 65,5);
    cell->setProperty("Value", ui->le_eqpt->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,5);
    cell->setProperty("Value", ui->le_job_no->text());

    cell = sheet->querySubObject("Cells(int,int)", 67,5);
    cell->setProperty("Value", ui->le_job_no->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,7);
    cell->setProperty("Value", ui->le_dr_no->text());

    cell = sheet->querySubObject("Cells(int,int)", 67,7);
    cell->setProperty("Value", ui->le_dr_no->text());

    cell = sheet->querySubObject("Cells(int,int)", 6,5);
    cell->setProperty("Value", ui->le_doc_no->text());
    cell = sheet->querySubObject("Cells(int,int)", 68,5);
    cell->setProperty("Value", ui->le_doc_no->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,11);
    cell->setProperty("Value", ui->le_rev->text());

    cell = sheet->querySubObject("Cells(int,int)", 67,11);
    cell->setProperty("Value", ui->le_rev->text());

    cell = sheet->querySubObject("Cells(int,int)", 3,11);
    cell->setProperty("Value", "1");
    cell = sheet->querySubObject("Cells(int,int)", 65,11);
    cell->setProperty("Value", "2");
    workbook->dynamicCall("Save()");
    workbook->dynamicCall("ExportAsFixedFormat(int, const QString&, int, BOOL, BOOL)", 0, fileName_pdf, 0, false, false);
    workbook->dynamicCall("Close()");
    workbook->dynamicCall("Quit()");
    excel->dynamicCall("Quit()");
    file_write.close();


    QMessageBox msgBox;
    msgBox.setText("Report has been generated... ");
    msgBox.exec();
}

void MainWindow::on_comboBox_activated(int index)
{
    ui->stackedWidget->setCurrentIndex(index);
}


void MainWindow::on_comboBox_currentIndexChanged(const QString &arg1)
{
    QString temp = arg1.trimmed();
    if(temp == "LIFTING LUG SHELL COVER/CHANNEL") {
        ui->stackedWidget->setCurrentIndex(0);
    }
    else if (temp == "LIFTING LUG FOR FLOATING HEAD") {
         ui->stackedWidget->setCurrentIndex(1);
    }
    else if (temp == "LIFTING LUG CHANNEL COVER") {
         ui->stackedWidget->setCurrentIndex(2);
    }
}

void MainWindow::on_fh_pb_check_for_shear_clicked()
{
    double allowable_shear = 0.4 * ui->fh_le_sy->text().toDouble();
    ui->fh_allowable_shear->setText(QString::number(allowable_shear, 'g', 3));
    double H22 = ui->fh_le_wt_to_be_lifted->text().toDouble();
    double H23 = ui->fh_le_shock_factor->text().toDouble();
    double H24 = ui->fh_le_no_of_liftin_lug->text().toDouble();
    double H36 = ui->fh_le_radius->text().toDouble();
    double H35 = ui->fh_le_diameter_of_hole->text().toDouble();


    double thickness_of_lifting_lug = ((H22*H23)/(H24*2*(H36-(H35/2))*allowable_shear));
    ui->fh_shear_thickness_of_lug->setText(QString::number(thickness_of_lifting_lug, 'g', 3));
    ui->fh_le_provided_thickness->setText(ui->fh_le_thickness_of_lug->text());

    if( ui->fh_le_provided_thickness->text().toDouble() > ui->fh_shear_thickness_of_lug->text().toDouble() ) {
        ui->fh_cb_safe->setChecked(true);
        ui->fh_cb_safe->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->fh_cb_safe->setChecked(false);
        ui->fh_cb_safe->setStyleSheet("background-color: rgb(255, 0, 0);");
    }
}

void MainWindow::on_fh_pb_check_for_shear_2_clicked()
{
    double allowable_shear = 0.4 * ui->fh_le_sy_2->text().toDouble();
    ui->fh_allowable_shear_2->setText(QString::number(allowable_shear, 'g', 3));
    double H22 = ui->fh_le_wt_to_be_lifted_2->text().toDouble();
    double H23 = ui->fh_le_shock_factor_2->text().toDouble();
    double H24 = ui->fh_le_no_of_liftin_lug_2->text().toDouble();
    double H36 = ui->fh_le_radius_2->text().toDouble();
    double H35 = ui->fh_le_diameter_of_hole_2->text().toDouble();


    double thickness_of_lifting_lug = ((H22*H23)/(H24*2*(H36-(H35/2))*allowable_shear));
    ui->fh_shear_thickness_of_lug_2->setText(QString::number(thickness_of_lifting_lug, 'g', 3));
    ui->fh_le_provided_thickness_2->setText(ui->fh_le_thickness_of_lug_2->text());

    if( ui->fh_le_provided_thickness_2->text().toDouble() > ui->fh_shear_thickness_of_lug_2->text().toDouble() ) {
        ui->fh_cb_safe_2->setChecked(true);
        ui->fh_cb_safe_2->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->fh_cb_safe_2->setChecked(false);
        ui->fh_cb_safe_2->setStyleSheet("background-color: rgb(255, 0, 0);");
    }
}


void MainWindow::on_fh_pb_check_for_bending_clicked()
{
    double H22 = ui->fh_le_wt_to_be_lifted->text().toDouble();
    double H23 = ui->fh_le_shock_factor->text().toDouble();
    double H32 = ui->fh_le_distance_of_lifting_lug->text().toDouble();
    double H24 = ui->fh_le_no_of_liftin_lug->text().toDouble();
    double H34 = ui->fh_le_thickness_of_lug->text().toDouble();
    double H33 = ui->fh_le_width_of_log->text().toDouble();
    double induced_bending_stress = (6*H22*H23*H32)/(H24*(H34*H34)*H33);
    ui->fh_le_induced_bending_stress->setText(QString::number(induced_bending_stress, 'g', 3));
    double allowable_bending_stress = 0.66 * ui->fh_le_sy->text().toDouble();
    ui->fh_le_bending_stree_of_Log->setText(QString::number(allowable_bending_stress, 'g', 3));

    if( allowable_bending_stress > induced_bending_stress ) {
        ui->fh_cb_safe_bending->setChecked(true);
        ui->fh_cb_safe_bending->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->fh_cb_safe_bending->setChecked(false);
        ui->fh_cb_safe_bending->setStyleSheet("background-color: rgb(255, 0, 0);");
    }

}

void MainWindow::on_fh_pb_check_for_bending_2_clicked()
{
    double H22 = ui->fh_le_wt_to_be_lifted_2->text().toDouble();
    double H23 = ui->fh_le_shock_factor_2->text().toDouble();
    double H32 = ui->fh_le_distance_of_lifting_lug_2->text().toDouble();
    double H24 = ui->fh_le_no_of_liftin_lug_2->text().toDouble();
    double H34 = ui->fh_le_thickness_of_lug_2->text().toDouble();
    double H33 = ui->fh_le_width_of_log_2->text().toDouble();
    double induced_bending_stress = (6*H22*H23*H32)/(H24*(H34*H34)*H33);
    ui->fh_le_induced_bending_stress_2->setText(QString::number(induced_bending_stress, 'g', 3));
    double allowable_bending_stress = 0.66 * ui->fh_le_sy_2->text().toDouble();
    ui->fh_le_bending_stree_of_Log_2->setText(QString::number(allowable_bending_stress, 'g', 3));

    if( allowable_bending_stress > induced_bending_stress ) {
        ui->fh_cb_safe_bending_2->setChecked(true);
        ui->fh_cb_safe_bending_2->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->fh_cb_safe_bending_2->setChecked(false);
        ui->fh_cb_safe_bending_2->setStyleSheet("background-color: rgb(255, 0, 0);");
    }

}



void MainWindow::on_fh_pb_check_for_shear_in_weld_2_clicked()
{
    double H33 = ui->fh_le_width_of_log_2->text().toDouble();
    double H34 = ui->fh_le_thickness_of_lug_2->text().toDouble();

    ui->fh_le_total_weld_length_2->setText(QString::number(H33, 'g', 3));
    ui->fh_le_total_weld_width_2->setText(QString::number(H34, 'g', 3));

    float total_weld_length = 2*(H33+H34);
    ui->fh_le_weld_length_2->setText(QString::number(total_weld_length, 'g', 3));

    double filet_weld_size = ui->fh_le_fillet_weld_leg_size_f_2->text().toDouble();
    ui->fh_le_fillet_weld_size_2->setText(QString::number(filet_weld_size, 'g', 3));

    double weld_area = total_weld_length * filet_weld_size * 0.707;
    ui->fh_le_weld_area_2->setText(QString::number(weld_area, 'f'));

    double H22 = ui->fh_le_wt_to_be_lifted_2->text().toDouble();
    double H23 = ui->fh_le_shock_factor_2->text().toDouble();
    double H24 = ui->fh_le_no_of_liftin_lug_2->text().toDouble();
    double induced_shear_stress = (H22*H23/H24)/weld_area;
    ui->fh_le_induce_shear_stress_in_weld_2->setText(QString::number(induced_shear_stress, 'g', 3));

    double allowable_eq_stress =0.4 * ui->fh_le_sy_2->text().toDouble();


    ui->fh_le_allowable_equivalet_stress_Sw_2->setText(QString::number(allowable_eq_stress, 'g', 3));

    if( allowable_eq_stress >induced_shear_stress) {
        ui->fh_cb_safe_shear_in_weld_2->setChecked(true);
        ui->fh_cb_safe_shear_in_weld_2->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->fh_cb_safe_shear_in_weld_2->setChecked(false);
        ui->fh_cb_safe_shear_in_weld_2->setStyleSheet("background-color: rgb(255, 0, 0);");
    }


}

void MainWindow::on_fh_pb_check_for_shear_in_weld_clicked()
{
    double H33 = ui->fh_le_width_of_log->text().toDouble();
    double H34 = ui->fh_le_thickness_of_lug->text().toDouble();

    ui->fh_le_total_weld_length->setText(QString::number(H33, 'g', 3));
    ui->fh_le_total_weld_width->setText(QString::number(H34, 'g', 3));

    float total_weld_length = 2*(H33+H34);
    ui->fh_le_weld_length->setText(QString::number(total_weld_length, 'g', 3));

    double filet_weld_size = ui->fh_le_fillet_weld_leg_size_f->text().toDouble();
    ui->fh_le_fillet_weld_size->setText(QString::number(filet_weld_size, 'g', 3));

    double weld_area = total_weld_length * filet_weld_size * 0.707;
    ui->fh_le_weld_area->setText(QString::number(weld_area, 'f'));

    double H22 = ui->fh_le_wt_to_be_lifted->text().toDouble();
    double H23 = ui->fh_le_shock_factor->text().toDouble();
    double H24 = ui->fh_le_no_of_liftin_lug->text().toDouble();
    double induced_shear_stress = (H22*H23/H24)/weld_area;
    ui->fh_le_induce_shear_stress_in_weld->setText(QString::number(induced_shear_stress, 'g', 3));

    double allowable_eq_stress =0.4 * ui->fh_le_sy->text().toDouble();


    ui->fh_le_allowable_equivalet_stress_Sw->setText(QString::number(allowable_eq_stress, 'g', 3));

    if( allowable_eq_stress >induced_shear_stress) {
        ui->fh_cb_safe_shear_in_weld->setChecked(true);
        ui->fh_cb_safe_shear_in_weld->setStyleSheet("background-color: rgb(0, 255, 0);");
    }
    else {
        ui->fh_cb_safe_shear_in_weld->setChecked(false);
        ui->fh_cb_safe_shear_in_weld->setStyleSheet("background-color: rgb(255, 0, 0);");
    }


}


void MainWindow::on_fh_pb_generate_report_clicked()
{
    QString fileName = "C:/ISGEC-TOOLS/Lifting Lug for floating head.xlsx";
    QString fileName_pdf = "C:/ISGEC-TOOLS/Lifting_Lug_floating_head";
    QFile file_write(fileName);
    if(file_write.open(QIODevice::ReadWrite)) {
       excel     = new QAxObject("Excel.Application");
       workbooks = excel->querySubObject("Workbooks");
       workbook  = workbooks->querySubObject("Open(const QString&)",fileName);
       sheets    = workbook->querySubObject("Worksheets");
       sheet     = sheets->querySubObject("Item(int)", 1);
    }
    else {
        QMessageBox msgBox;
        msgBox.setText("Please Make Sure Excel file is closed!!");
        msgBox.exec();
        return;;
    }
    file_write.close();
    // on_pb_check_for_shear_clicked ////////////////////////////////////////////////////////////////
/*    if(ui->fh_le_provided_thickness->text().toDouble() > ui->fh_shear_thickness_of_lug->text().toDouble() ) {
        auto cell = sheet->querySubObject("Cells(int,int)", 56,8);
        cell->setProperty("Value", "SAFE");
    }
    else {
        auto cell = sheet->querySubObject("Cells(int,int)", 56,8);
        cell->setProperty("Value", "UNSAFE");
    }*/

    auto cell = sheet->querySubObject("Cells(int,int)", 5,3);
    cell->setProperty("Value", ui->fh_le_designed_by->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,3);
    cell->setProperty("Value", ui->fh_le_designed_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 1,5);
    cell->setProperty("Value", ui->fh_le_client->text());
    cell = sheet->querySubObject("Cells(int,int)", 52,5);
    cell->setProperty("Value", ui->fh_le_client->text());

    cell = sheet->querySubObject("Cells(int,int)", 3,5);
    cell->setProperty("Value", ui->fh_le_eqpt->text());
    cell = sheet->querySubObject("Cells(int,int)", 54,5);
    cell->setProperty("Value", ui->fh_le_eqpt->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,5);
    cell->setProperty("Value", ui->fh_le_job_no->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,5);
    cell->setProperty("Value", ui->fh_le_job_no->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,7);
    cell->setProperty("Value", ui->fh_le_dr_no->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,7);
    cell->setProperty("Value", ui->fh_le_dr_no->text());

    cell = sheet->querySubObject("Cells(int,int)", 6,5);
    cell->setProperty("Value", ui->fh_le_doc_no->text());
    cell = sheet->querySubObject("Cells(int,int)", 57,5);
    cell->setProperty("Value", ui->fh_le_doc_no->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,11);
    cell->setProperty("Value", ui->fh_le_rev->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,11);
    cell->setProperty("Value", ui->fh_le_rev->text());

    cell = sheet->querySubObject("Cells(int,int)", 3,11);
    cell->setProperty("Value", "1");
    cell = sheet->querySubObject("Cells(int,int)", 54,11);
    cell->setProperty("Value", "2");

    cell = sheet->querySubObject("Cells(int,int)", 22,8);
    cell->setProperty("Value", ui->fh_le_wt_to_be_lifted->text());

    cell = sheet->querySubObject("Cells(int,int)", 23,8);
    cell->setProperty("Value", ui->fh_le_shock_factor->text());

    cell = sheet->querySubObject("Cells(int,int)", 24,8);
    cell->setProperty("Value", ui->fh_le_no_of_liftin_lug->text());

    cell = sheet->querySubObject("Cells(int,int)", 27,8);
    cell->setProperty("Value", ui->fh_cb_lifting_ug_materia->currentText());
    cell = sheet->querySubObject("Cells(int,int)", 27,9);
    cell->setProperty("Value", ui->fh_cb_grade->currentText());

    cell = sheet->querySubObject("Cells(int,int)", 28,8);
    cell->setProperty("Value", ui->fh_le_sy->text());

    cell = sheet->querySubObject("Cells(int,int)", 32,8);
    cell->setProperty("Value", ui->fh_le_distance_of_lifting_lug->text());

    cell = sheet->querySubObject("Cells(int,int)", 33,8);
    cell->setProperty("Value", ui->fh_le_width_of_log->text());

    cell = sheet->querySubObject("Cells(int,int)", 34,8);
    cell->setProperty("Value", ui->fh_le_thickness_of_lug->text());

    cell = sheet->querySubObject("Cells(int,int)", 35,8);
    cell->setProperty("Value", ui->fh_le_diameter_of_hole->text());

    cell = sheet->querySubObject("Cells(int,int)", 36,8);
    cell->setProperty("Value", ui->fh_le_radius->text());

    cell = sheet->querySubObject("Cells(int,int)", 37,8);
    cell->setProperty("Value", ui->fh_le_fillet_weld_leg_size_f->text());

    cell = sheet->querySubObject("Cells(int,int)", 41,8);
    cell->setProperty("Value", ui->fh_allowable_shear->text());

    cell = sheet->querySubObject("Cells(int,int)", 43,8);
    cell->setProperty("Value", ui->fh_shear_thickness_of_lug->text());

    cell = sheet->querySubObject("Cells(int,int)", 45,5);
    cell->setProperty("Value", ui->fh_le_thickness_of_lug->text());

    cell = sheet->querySubObject("Cells(int,int)", 45,7);
    cell->setProperty("Value", ui->fh_shear_thickness_of_lug->text());

    cell = sheet->querySubObject("Cells(int,int)", 49,8);
    cell->setProperty("Value", ui->fh_le_induced_bending_stress->text());
    cell = sheet->querySubObject("Cells(int,int)", 58,8);
    cell->setProperty("Value", ui->fh_le_bending_stree_of_Log->text());

    cell = sheet->querySubObject("Cells(int,int)", 64,8);
    cell->setProperty("Value", ui->fh_le_total_weld_length->text());

    cell = sheet->querySubObject("Cells(int,int)", 65,8);
    cell->setProperty("Value", ui->fh_le_total_weld_width->text());

    cell = sheet->querySubObject("Cells(int,int)", 66,8);
    cell->setProperty("Value", ui->fh_le_weld_length->text());

    cell = sheet->querySubObject("Cells(int,int)", 67,8);
    cell->setProperty("Value", ui->fh_le_fillet_weld_size->text());

    cell = sheet->querySubObject("Cells(int,int)", 68,8);
    cell->setProperty("Value", ui->fh_le_weld_area->text());

    cell = sheet->querySubObject("Cells(int,int)", 70,8);
    cell->setProperty("Value", ui->fh_le_induce_shear_stress_in_weld->text());

    cell = sheet->querySubObject("Cells(int,int)", 72,8);
    cell->setProperty("Value", ui->fh_le_allowable_equivalet_stress_Sw->text());



    workbook->dynamicCall("Save()");
    workbook->dynamicCall("ExportAsFixedFormat(int, const QString&, int, BOOL, BOOL)", 0, fileName_pdf, 0, false, false);
    workbook->dynamicCall("Close()");
    workbook->dynamicCall("Quit()");
    excel->dynamicCall("Quit()");
    file_write.close();

    QMessageBox msgBox;
    msgBox.setText("Report has been generated... ");
    msgBox.exec();
}

void MainWindow::on_fh_pb_generate_report_2_clicked()
{
    QString fileName = "C:/ISGEC-TOOLS/Lifting Lug for CHANNEL COVER.xlsx";
    QString fileName_pdf = "C:/ISGEC-TOOLS/Lug_for_CHANNEL_COVER";
    QFile file_write(fileName);
    if(file_write.open(QIODevice::ReadWrite)) {
       excel     = new QAxObject("Excel.Application");
       workbooks = excel->querySubObject("Workbooks");
       workbook  = workbooks->querySubObject("Open(const QString&)",fileName);
       sheets    = workbook->querySubObject("Worksheets");
       sheet     = sheets->querySubObject("Item(int)", 1);
    }
    else {
        QMessageBox msgBox;
        msgBox.setText("Please Make Sure Excel file is closed!!");
        msgBox.exec();
        return;;
    }
    file_write.close();
    auto cell = sheet->querySubObject("Cells(int,int)", 5,3);
    cell->setProperty("Value", ui->fh_le_designed_by_2->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,3);
    cell->setProperty("Value", ui->fh_le_designed_by_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 1,5);
    cell->setProperty("Value", ui->fh_le_client_2->text());
    cell = sheet->querySubObject("Cells(int,int)", 52,5);
    cell->setProperty("Value", ui->fh_le_client_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 3,5);
    cell->setProperty("Value", ui->fh_le_eqpt_2->text());
    cell = sheet->querySubObject("Cells(int,int)", 54,5);
    cell->setProperty("Value", ui->fh_le_eqpt_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,5);
    cell->setProperty("Value", ui->fh_le_job_no_2->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,5);
    cell->setProperty("Value", ui->fh_le_job_no_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,7);
    cell->setProperty("Value", ui->fh_le_dr_no_2->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,7);
    cell->setProperty("Value", ui->fh_le_dr_no_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 6,5);
    cell->setProperty("Value", ui->fh_le_doc_no_2->text());
    cell = sheet->querySubObject("Cells(int,int)", 57,5);
    cell->setProperty("Value", ui->fh_le_doc_no_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,11);
    cell->setProperty("Value", ui->fh_le_rev_2->text());
    cell = sheet->querySubObject("Cells(int,int)", 56,11);
    cell->setProperty("Value", ui->fh_le_rev_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 3,11);
    cell->setProperty("Value", "1");
    cell = sheet->querySubObject("Cells(int,int)", 54,11);
    cell->setProperty("Value", "2");

    cell = sheet->querySubObject("Cells(int,int)", 22,8);
    cell->setProperty("Value", ui->fh_le_wt_to_be_lifted_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 23,8);
    cell->setProperty("Value", ui->fh_le_shock_factor_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 24,8);
    cell->setProperty("Value", ui->fh_le_no_of_liftin_lug_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 27,8);
    cell->setProperty("Value", ui->fh_cb_lifting_ug_materia_2->currentText());
    cell = sheet->querySubObject("Cells(int,int)", 27,9);
    cell->setProperty("Value", ui->fh_cb_grade_2->currentText());

    cell = sheet->querySubObject("Cells(int,int)", 28,8);
    cell->setProperty("Value", ui->fh_le_sy_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 32,8);
    cell->setProperty("Value", ui->fh_le_distance_of_lifting_lug_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 33,8);
    cell->setProperty("Value", ui->fh_le_width_of_log_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 34,8);
    cell->setProperty("Value", ui->fh_le_thickness_of_lug_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 35,8);
    cell->setProperty("Value", ui->fh_le_diameter_of_hole_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 36,8);
    cell->setProperty("Value", ui->fh_le_radius_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 37,8);
    cell->setProperty("Value", ui->fh_le_fillet_weld_leg_size_f_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 41,8);
    cell->setProperty("Value", ui->fh_allowable_shear_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 43,8);
    cell->setProperty("Value", ui->fh_shear_thickness_of_lug_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 45,5);
    cell->setProperty("Value", ui->fh_le_thickness_of_lug_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 45,7);
    cell->setProperty("Value", ui->fh_shear_thickness_of_lug_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 49,8);
    cell->setProperty("Value", ui->fh_le_induced_bending_stress_2->text());
    cell = sheet->querySubObject("Cells(int,int)", 58,8);
    cell->setProperty("Value", ui->fh_le_bending_stree_of_Log_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 64,8);
    cell->setProperty("Value", ui->fh_le_total_weld_length_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 65,8);
    cell->setProperty("Value", ui->fh_le_total_weld_width_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 66,8);
    cell->setProperty("Value", ui->fh_le_weld_length_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 67,8);
    cell->setProperty("Value", ui->fh_le_fillet_weld_size_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 68,8);
    cell->setProperty("Value", ui->fh_le_weld_area_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 70,8);
    cell->setProperty("Value", ui->fh_le_induce_shear_stress_in_weld_2->text());

    cell = sheet->querySubObject("Cells(int,int)", 72,8);
    cell->setProperty("Value", ui->fh_le_allowable_equivalet_stress_Sw_2->text());



    workbook->dynamicCall("Save()");
    workbook->dynamicCall("ExportAsFixedFormat(int, const QString&, int, BOOL, BOOL)", 0, fileName_pdf, 0, false, false);
    workbook->dynamicCall("Close()");
    workbook->dynamicCall("Quit()");
    excel->dynamicCall("Quit()");
    file_write.close();

    QMessageBox msgBox;
    msgBox.setText("Report has been generated... ");
    msgBox.exec();

}


void MainWindow::on_comboBox_currentIndexChanged(int index)
{

}

void MainWindow::on_comboBox_activated(const QString &arg1)
{
    QString temp = arg1.trimmed();
    if(temp == "LIFTING LUG SHELL COVER/CHANNEL") {
        ui->stackedWidget->setCurrentIndex(0);
    }
    else if (temp == "LIFTING LUG FOR FLOATING HEAD") {
         ui->stackedWidget->setCurrentIndex(1);
    }
    else if (temp == "LIFTING LUG CHANNEL COVER") {
         ui->stackedWidget->setCurrentIndex(2);
    }

}

void MainWindow::on_fh_cb_grade_2_activated(const QString &arg1)
{

}


void MainWindow::on_le_thickness_of_lug_editingFinished()
{
    int fillet_value = 0.7 * ui->le_thickness_of_lug->text().toInt();
    ui->le_fillet_lug_to_pad->setText(QString::number(fillet_value +1));
}


void MainWindow::on_le_thickness_of_pad_editingFinished()
{
    int fillet_value = 0.7 * ui->le_thickness_of_pad->text().toInt();
    ui->le_fillet_weld_leg_size_f2->setText(QString::number(fillet_value +1));
}


void MainWindow::on_fh_le_thickness_of_lug_editingFinished()
{
    int fillet_value = 0.7 * ui->fh_le_thickness_of_lug->text().toInt();
    ui->fh_le_fillet_weld_leg_size_f->setText(QString::number(fillet_value +1));
}


void MainWindow::on_fh_le_thickness_of_lug_2_editingFinished()
{
    int fillet_value = 0.7 * ui->fh_le_thickness_of_lug_2->text().toInt();
    ui->fh_le_fillet_weld_leg_size_f_2->setText(QString::number(fillet_value +1));
}

