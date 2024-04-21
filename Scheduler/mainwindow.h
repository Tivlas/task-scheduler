#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "src/pages/pages.h"

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_createButton_clicked();

    void on_refreshButton_clicked();

private:
    Ui::MainWindow *ui;
    pages* taskCreationPages;

    void showTasks();
};
#endif // MAINWINDOW_H
