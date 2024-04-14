#ifndef PAGES_H
#define PAGES_H

#include <QWidget>
#include <QDialog>

namespace Ui {
class pages;
}

class pages : public QDialog
{
    Q_OBJECT

public:
    explicit pages(QWidget *parent = nullptr);
    ~pages();
    const int DONE_CODE = 0;
    const int CANCELED_CODE = 1;

private slots:
    void on_prevButton_clicked();

    void on_nextButton_clicked();

    void on_addTaskButton_clicked();

    void on_cancelButton_clicked();

    void on_onceRadioButton_clicked();

    void on_dailyRadioButton_clicked();

    void on_weeklyRadioButton_clicked();

    void on_chooseFileButton_clicked();

signals:
    void cancelSignal();
    void doneSignal();

private:
    Ui::pages *ui;
};

#endif // PAGES_H
