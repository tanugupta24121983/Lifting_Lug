#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFile>
#include <QAxObject>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();


    void on_cb_thickness_activated_custom(const QString &arg1);
    void on_fh_cb_thickness_activated_custom(const QString &arg1);
    void on_cb_grade_activated_custom(const QString &arg1);
    void on_fh_cb_grade_activated_custom(const QString &arg1);
    void on_lc_cb_grade_activated_custom(const QString &arg1);
    void load_temprature_data();
private slots:
    void on_pb_check_for_shear_clicked();

    void on_le_wt_to_be_lifted_editingFinished();

    void on_le_shock_factor_editingFinished();

    void on_le_no_of_liftin_lugs_editingFinished();

    void on_cb_lifting_ug_materia_activated(const QString &arg1);
    void on_cb_thickness_activated(const QString &arg1);

    void on_cb_grade_activated(const QString &arg1);

    void on_le_temprature_editingFinished();

    void on_pb_check_for_bending_clicked();

    void on_pb_check_for_shear_in_weld_clicked();

    void on_pb_generate_report_clicked();

    void saveSettings();
    void loadSettings();
    void on_comboBox_activated(int index);

    void on_comboBox_currentIndexChanged(const QString &arg1);

    void on_fh_cb_lifting_ug_materia_activated(const QString &arg1);

    void on_fh_cb_grade_activated(const QString &arg1);

    void on_fh_cb_thicknes_activated(const QString &arg1);
    void on_fh_cb_2_thickness_activated_custom(const QString &arg1);

    void on_fh_le_tempratur_editingFinished();

    void on_fh_pb_check_for_shear_clicked();

    void on_fh_pb_check_for_bending_clicked();

    void on_fh_pb_check_for_shear_in_weld_clicked();

    void on_fh_pb_generate_report_clicked();

    void on_comboBox_currentIndexChanged(int index);

    void on_comboBox_activated(const QString &arg1);

    void on_fh_cb_grade_2_textActivated(const QString &arg1);


    void on_fh_cb_thicknes_2_activated(const QString &arg1);
    void on_fh_cb_grade_2_activated_custom(const QString &arg1);

    void on_fh_cb_lifting_ug_materia_2_textActivated(const QString &arg1);

    void on_fh_le_tempratur_2_editingFinished();

    void on_fh_pb_check_for_shear_2_clicked();

    void on_fh_pb_check_for_bending_2_clicked();

    void on_fh_pb_check_for_shear_in_weld_2_clicked();

    void on_fh_cb_grade_2_activated(const QString &arg1);

    void on_fh_pb_generate_report_2_clicked();

    void on_le_thickness_of_lug_editingFinished();

    void on_le_thickness_of_pad_editingFinished();

    void on_fh_le_thickness_of_lug_editingFinished();

    void on_fh_le_thickness_of_lug_2_editingFinished();

private :
    Ui::MainWindow *ui;
    QAxObject * excel;
    QAxObject * workbooks;
    QAxObject * workbook;
    QAxObject * sheets;
    QAxObject * sheet;


};
#endif // MAINWINDOW_H
