#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFileDialog>
#include <QAxObject>
#include <QDebug>
#include <QTableWidget>
#include <QResizeEvent>
#include <QVector>
#include <QMessageBox>

#include "calculate.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
    void setColWidth(QTableWidget*, QTableWidget*); // подгон размеров таблиц
    void enabledButtons();                          // при старте заблокированы кнопка анализа текста, а также выгрузки и вызова *excel-файла
    inline void TabBaseSettExcel();
    inline void TabBaseSettFiles();

protected:
    void resizeEvent(QResizeEvent *);

signals:
    void updFileNames();

public slots:
    void updateFileNames(); // обновить список загруженных документов

private slots:
    void on_pushButton_exit_clicked();

    void on_pushButton_open_doc_clicked();

    void on_pushButton_analize_clicked();

    void on_pushButton_openExel_clicked();

    void on_pushButton_toExcel_clicked();

private:
    Ui::MainWindow *ui;

    QStringList fname; // имена загруженных файлов

    QString textResultAll; // весь текст
    QString textResult;    // текст одного загруженного документа

    QAxObject* wordApplication; // объекты
    QAxObject* documents;       // для
    QAxObject* document;        // загрузки
    QAxObject* words;           // .doc(x)-файлов

    QTableWidget* tabWidExcel;       // TableWidget с аналитическими данными
    QTableWidget* tabWidFiles;       // TableWidget с именами загруженных файлов
    QStringList namesFirstCol_excel; // названия анализируемых параметров, выводятся в первом столбце
    QStringList tabHeader_files;     // заголовок для загруженных файлов
    QStringList tabHeader_excel;     // заголовки для аналитических даннымх

    QString excelName; // имя файла для выгрузки в Excel

    QVector <int> outputVal; // аналитические параметры

    int RowAnalize;   // строк в таблице с аналитикой
    int RowFiles;     // строк в TabWidget со списком загруженных файлов
    int ColAnalize;   // столбцов в таблице с аналитикой
    int ColFiles;     // столбцов в TabWidget со списком загруженных файлов

    QHeaderView *verticalHeader;

    calculate* calc;

};

#endif // MAINWINDOW_H
