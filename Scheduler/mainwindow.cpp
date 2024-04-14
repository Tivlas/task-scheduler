#include "mainwindow.h"
#include "./ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_prevButton_clicked()
{
    int newIdx = ui->pages->currentIndex() - 1;
    ui->pages->setCurrentIndex(newIdx);
    if (newIdx == ui->pages->count() - 1) {
        ui->nextButton->setText("Далее");
    }
}


void MainWindow::on_nextButton_clicked()
{
    int newIdx = ui->pages->currentIndex() + 1;
    ui->pages->setCurrentIndex(newIdx);
    if (newIdx == ui->pages->count() - 1) {
        ui->nextButton->setText("Готово");
    }
}


void MainWindow::on_onceRadioButton_clicked()
{
    ui->weeklyFrame->hide();
    ui->dailyFrame->hide();
}


void MainWindow::on_dailyRadioButton_clicked()
{
    ui->weeklyFrame->hide();
    ui->dailyFrame->show();
}


void MainWindow::on_weeklyRadioButton_clicked()
{
    ui->weeklyFrame->show();
    ui->dailyFrame->hide();
}

