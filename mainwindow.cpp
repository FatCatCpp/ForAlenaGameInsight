#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    this->setWindowTitle("Анализ файлов Microsoft Office Word");
    this->setFixedSize(this->width(), this->height());
    excelName = "D:\\Анализ_текстов.xlsx";

    calc = new calculate();

    RowAnalize = 8;
    RowFiles = 200;
    ColAnalize = 2;
    ColFiles = 1;

    ui->pushButton_analize->setDisabled(true);
    ui->pushButton_openExel->setDisabled(true);
    ui->pushButton_toExcel->setDisabled(true);

    TabBaseSettFiles();
    TabBaseSettExcel();

    connect(this, SIGNAL(updFileNames()), this, SLOT(updateFileNames()));
}

inline void MainWindow::TabBaseSettExcel()
{
    tabWidExcel = new QTableWidget(ui->tableWidget);
    tabWidExcel->setRowCount(RowAnalize);
    tabWidExcel->setColumnCount(ColAnalize);
    namesFirstCol_excel << "Количествово файлов" << "Размер файлов, кб" << "Количествово слов" << "Количествово символов"
                        << "Количествово цифр" << "Количествово знаков препинания" << "Количествово пробелов" << "Слово \"кабель\" в тексте";
    tabHeader_excel << "Параметр" << "Значение";
    tabWidExcel->setHorizontalHeaderLabels(tabHeader_excel);

    for (int i = 0; i < RowAnalize; i++)
        tabWidExcel->setItem(i, 0, new QTableWidgetItem(namesFirstCol_excel[i]));

    tabWidExcel->verticalHeader()->setVisible(false);
    tabWidExcel->setEditTriggers(QAbstractItemView::NoEditTriggers);
    tabWidExcel->setSelectionBehavior(QAbstractItemView::SelectRows);
    tabWidExcel->setSelectionMode(QAbstractItemView::SingleSelection);
    tabWidExcel->setShowGrid(false);
    verticalHeader = tabWidExcel->verticalHeader();
    verticalHeader->setSectionResizeMode(QHeaderView::Fixed);
    verticalHeader->setDefaultSectionSize(20);
}

inline void MainWindow::TabBaseSettFiles()
{
    tabWidFiles = new QTableWidget(ui->tableWidget_filesLoad);
    tabWidFiles->setColumnCount(ColFiles);
    tabWidFiles->setRowCount(RowFiles);
    tabHeader_files << "Загруженные файлы";
    tabWidFiles->setHorizontalHeaderLabels(tabHeader_files);
    tabWidFiles->verticalHeader()->setVisible(false);
    tabWidFiles->setEditTriggers(QAbstractItemView::NoEditTriggers);
    tabWidFiles->setSelectionBehavior(QAbstractItemView::SelectRows);
    tabWidFiles->setSelectionMode(QAbstractItemView::SingleSelection);
    tabWidFiles->setShowGrid(false);
    verticalHeader = tabWidFiles->verticalHeader();
    verticalHeader->setSectionResizeMode(QHeaderView::Fixed);
    verticalHeader->setDefaultSectionSize(20);
}

void MainWindow::resizeEvent(QResizeEvent *e)
{
    QMainWindow::resizeEvent(e);

    tabWidFiles->resize(e->size().width(), e->size().height());
    tabWidExcel->resize(e->size().width(), e->size().height());

    setColWidth(tabWidExcel, tabWidFiles);
}

void MainWindow::enabledButtons()
{
    ui->pushButton_analize->setDisabled(false);
    ui->pushButton_openExel->setDisabled(false);
    ui->pushButton_toExcel->setDisabled(false);
}

void MainWindow::setColWidth(QTableWidget* tab1, QTableWidget* tab2)
{
    tab1->setColumnWidth(0, ui->tableWidget->width()*0.7);
    tab1->setColumnWidth(1, ui->tableWidget->width()*0.3);

    tab2->setColumnWidth(0, ui->tableWidget_filesLoad->width());
}

MainWindow::~MainWindow()
{
    delete ui;

    delete wordApplication;
    delete documents;
    delete document;
    delete words;
    delete tabWidExcel;
    delete tabWidFiles;
    delete verticalHeader;
}

void MainWindow::on_pushButton_exit_clicked()
{
    exit(0);
}

void MainWindow::on_pushButton_open_doc_clicked()
{
    if ((!fname.isEmpty()) || (!fname.isEmpty())) {
        tabWidFiles->clear();
        fname.clear();
    }

    fname = QFileDialog::getOpenFileNames(this, tr("Открыть документ"), "D:\\Egor\\doc_files", "Docx files (*.docx);;Doc files (*.doc)");
    emit updFileNames();

    for (int i = 0; i < fname.size(); i++)  {
        wordApplication = new QAxObject("Word.Application", this);
        documents = wordApplication->querySubObject("Documents");
        document = documents->querySubObject("Open(const QString&, bool)", fname[i], true);
        words = document->querySubObject("Words");
        int countWord = words->dynamicCall("Count()").toInt();
        for (int a = 1; a <= countWord; a++) {
            textResult.append(words->querySubObject("Item(int)", a)->dynamicCall("Text()").toString());
        }
        wordApplication->dynamicCall("Quit()");
        textResultAll += textResult;
        textResult.clear();
    }

    enabledButtons();
}

void MainWindow::updateFileNames()
{
    tabWidFiles->clear();
    tabHeader_files << "Загруженные файлы";
    tabWidFiles->setHorizontalHeaderLabels(tabHeader_files);
    for (int i = 0; i < fname.size(); i++)
        tabWidFiles->setItem(i, 0, new QTableWidgetItem(fname[i]));
}

void MainWindow::on_pushButton_analize_clicked()
{
    QString strForCalc = textResultAll;

    if (!outputVal.isEmpty())
        outputVal.clear();

    outputVal.append(calc->CalcFilesCount(fname));
    outputVal.append(calc->CalcSizeFiles(fname));
    outputVal.append(calc->CalcWordCount(strForCalc));
    outputVal.append(calc->CalcSymbCount(strForCalc));
    outputVal.append(calc->CalcDigitCount(strForCalc));
    outputVal.append(calc->CalcPuncMarks(strForCalc));
    outputVal.append(calc->CalcSpaceCount(strForCalc));
    outputVal.append(calc->CalcConcreteWord(strForCalc, "кабель"));

    for (int i = 0; i < outputVal.size(); i++) {
        QTableWidgetItem* item = new QTableWidgetItem(QString::number(outputVal[i]));
        item->setTextAlignment(Qt::AlignCenter);
        tabWidExcel->setItem(i, 1, item);
    }
}

void MainWindow::on_pushButton_openExel_clicked()
{
    calc->startProcess(excelName);
}

void MainWindow::on_pushButton_toExcel_clicked()
{
    QAxObject *mExcel = new QAxObject("Excel.Application",this); // получаем указатель на Excel
    QAxObject *workbooks = mExcel->querySubObject("Workbooks"); // на книги
    QAxObject *workbook = workbooks->querySubObject( "Open(const QString&)", "D:\\Анализ_текстов.xlsx" ); // на директорию, откуда грузить книг
    QAxObject *mSheets = workbook->querySubObject( "Sheets" ); // на листы
    QAxObject *StatSheet = mSheets->querySubObject( "Item(const QVariant&)", QVariant("Лист1") ); // указываем, какой лист выбрать
    QAxObject* cell;

    for (int i = 2; i <= 2; i++) {
        for (int j = 1; j <= RowAnalize; j++) {
            cell = StatSheet->querySubObject("Cells(QVariant,QVariant)", i, j);
            cell->setProperty("Value", QVariant(outputVal[j-1]));
        }
    }

    if (!outputVal.isEmpty())
        outputVal.clear();

    // освобождение памяти
    delete cell;
    delete StatSheet;
    delete mSheets;
    workbook->dynamicCall("Save()");
    delete workbook;
    delete workbooks;
    mExcel->dynamicCall("Quit()");
    delete mExcel;

    QMessageBox msgBox;
    msgBox.setWindowTitle("Формирование аналитического файла");
    msgBox.setText("Запись данных в \"D:/Анализ_текстов.xlsx\" прошла успешно!");
    msgBox.exec();

}





