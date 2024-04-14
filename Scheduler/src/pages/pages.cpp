#include "pages.h"
#include "ui_pages.h"

pages::pages(QWidget *parent)
    : QDialog(parent)
    , ui(new Ui::pages)
{
    ui->setupUi(this);
    ui->addTaskButton->hide();
    ui->cancelButton->hide();
}

pages::~pages()
{
    delete ui;
}

void pages::on_prevButton_clicked()
{
    int newIdx = ui->createPages->currentIndex() - 1;
    ui->createPages->setCurrentIndex(newIdx);
    ui->nextButton->show();
    ui->addTaskButton->hide();
    ui->cancelButton->hide();
}


void pages::on_nextButton_clicked()
{
    int newIdx = ui->createPages->currentIndex() + 1;
    ui->createPages->setCurrentIndex(newIdx);
    if (newIdx == ui->createPages->count() - 1) {
        ui->nextButton->hide();
        ui->addTaskButton->show();
        ui->cancelButton->show();
    }
}


void pages::on_addTaskButton_clicked()
{
    // TODO add task
    this->done(DONE_CODE);
}


void pages::on_cancelButton_clicked()
{
    this->done(CANCELED_CODE);
}


void pages::on_onceRadioButton_clicked()
{
    ui->weeklyFrame->hide();
    ui->dailyFrame->hide();
}


void pages::on_dailyRadioButton_clicked()
{
    ui->weeklyFrame->hide();
    ui->dailyFrame->show();
}


void pages::on_weeklyRadioButton_clicked()
{
    ui->weeklyFrame->show();
    ui->dailyFrame->hide();
}


void pages::on_chooseFileButton_clicked()
{

}

