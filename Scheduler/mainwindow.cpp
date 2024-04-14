#include "mainwindow.h"
#include "./ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow), taskCreationPages(new pages())
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete taskCreationPages;
    delete ui;
}

void MainWindow::on_createButton_clicked()
{
    taskCreationPages->exec();
}
