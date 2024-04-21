#ifndef PAGES_H
#define PAGES_H

#include <QWidget>
#include <QDialog>
#include <QMessageBox>
#include <windows.h>
#include <comdef.h>
#include <wincred.h>
#include <taskschd.h>
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")

#define SUNDAY 1
#define MONDAY 2
#define TUESDAY 4
#define WEDNESDAY 8
#define THURSDAY 16
#define FRIDAY 32
#define SATURDAY 64

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
    enum ErrCode {
        Ok,
        Err
    };

    Ui::pages *ui;
    QString actionPath;

    ErrCode addSpecificTimeTask();
    ErrCode addDailyTask();
    ErrCode addWeeklyTask();

    QString getTaskName();
    QString getTaskDescription();
    QString getActionPath();
    QString getActionArgs();
    QDateTime getStartDateTime();
    void warningMsgBox(QString title, QString text);
    void errorMsgBox(QString title, QString text);
    void okMsgBox(QString title, QString text);


};

#endif // PAGES_H
