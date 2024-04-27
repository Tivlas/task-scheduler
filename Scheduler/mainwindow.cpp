#define _WIN32_DCOM
#include "mainwindow.h"
#include "./ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow), taskCreationPages(new pages())
{
    ui->setupUi(this);
    showTasks();
}

MainWindow::~MainWindow()
{
    delete taskCreationPages;
    delete ui;
}

void MainWindow::on_createButton_clicked()
{
    delete taskCreationPages;
    taskCreationPages = new pages();
    taskCreationPages->exec();
    showTasks();
}

void MainWindow::showTasks() {
    ui->taskListWidget->clear();
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed " << hr;
        return;
    }

    ITaskService *pService = NULL;
    hr = CoCreateInstance( CLSID_TaskScheduler,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ITaskService,
                          (void**)&pService );
    if (FAILED(hr))
    {
        qDebug() << "Failed to CoCreate an instance of the TaskService class";
        CoUninitialize();
        return;
    }

    //  Connect to the task service.
    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() <<"ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return;
    }

    //  ------------------------------------------------------
    //  Get the pointer to the root task folder.
    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder( _bstr_t( L"\\") , &pRootFolder );

    pService->Release();
    if( FAILED(hr) )
    {
        qDebug() <<"Cannot get Root Folder pointer";
        CoUninitialize();
        return;
    }

    //  -------------------------------------------------------
    //  Get the registered tasks in the folder.
    IRegisteredTaskCollection* pTaskCollection = NULL;
    hr = pRootFolder->GetTasks( NULL, &pTaskCollection );

    pRootFolder->Release();
    if( FAILED(hr) )
    {
        qDebug() <<"Cannot get the registered tasks";
        CoUninitialize();
        return;
    }

    LONG numTasks = 0;
    hr = pTaskCollection->get_Count(&numTasks);
    if (FAILED(hr)) {
        qDebug() << "Cannot get task count";
        return;
    }

    if( numTasks == 0 )
    {
        qDebug() <<"\nNo Tasks are currently running";
        pTaskCollection->Release();
        CoUninitialize();
        return;
    }

    for(LONG i=0; i < numTasks; i++)
    {
        IRegisteredTask* pRegisteredTask = NULL;
        hr = pTaskCollection->get_Item( _variant_t(i+1), &pRegisteredTask );

        if( SUCCEEDED(hr) )
        {
            BSTR taskName = NULL;
            hr = pRegisteredTask->get_Name(&taskName);
            if( SUCCEEDED(hr) )
            {
                ITaskDefinition *pTaskDef = NULL;
                hr = pRegisteredTask->get_Definition(&pTaskDef);
                if (SUCCEEDED(hr))
                {
                    IRegistrationInfo *pRegInfo = NULL;
                    hr = pTaskDef->get_RegistrationInfo(&pRegInfo);
                    if (SUCCEEDED(hr))
                    {
                        BSTR author = NULL;
                        hr = pRegInfo->get_Author(&author);
                        if (SUCCEEDED(hr))
                        {
                            QString authorStr = QString::fromWCharArray(author);
                            if (authorStr.toStdWString() == AUTHOR_NAME)
                            {
                                QString name = QString::fromWCharArray(taskName);
                                ui->taskListWidget->addItem(name);
                            }
                            SysFreeString(author);
                        }
                        else
                        {
                            qDebug() << "Cannot get the registered task author";
                        }
                        pRegInfo->Release();
                    }
                    else
                    {
                        qDebug() << "Cannot get the registered task registration info";
                    }
                    pTaskDef->Release();
                }
                else
                {
                    qDebug() << "Cannot get the registered task definition";
                }
                SysFreeString(taskName);
            }
            else
            {
                qDebug() << "Cannot get the registered task name";
            }
            pRegisteredTask->Release();
        }
        else
        {
            qDebug() << "Cannot get the registered task item at index";
        }
    }

    pTaskCollection->Release();
    CoUninitialize();
    return;
}

void MainWindow::on_refreshButton_clicked()
{
    ui->taskListWidget->clear();
    showTasks();
}


void MainWindow::on_taskListWidget_itemClicked(QListWidgetItem *item)
{
    QString text;
    QString name = item->text();
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed";
        return;
    }

    //  ------------------------------------------------------
    //  Create a name for the task.
    auto wname = name.toStdWString();
    LPCWSTR wszTaskName = wname.c_str();

    //  ------------------------------------------------------
    //  Create an instance of the Task Service.
    ITaskService *pService = NULL;
    hr = CoCreateInstance( CLSID_TaskScheduler,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ITaskService,
                          (void**)&pService );
    if (FAILED(hr))
    {
        qDebug() << "Failed to create an instance of ITaskService";
        CoUninitialize();
        return;
    }

    //  Connect to the task service.
    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() << "ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return;
    }

    //  ------------------------------------------------------
    //  Get the pointer to the root task folder.  This folder will hold the
    //  new task that is registered.
    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get Root Folder pointer";
        pService->Release();
        CoUninitialize();
        return;
    }

    // If the same task exists, remove it.
    IRegisteredTask *pRegisteredTask = NULL;
    hr = pRootFolder->GetTask(_bstr_t( wszTaskName), &pRegisteredTask);
    if (FAILED(hr))
    {
        pService->Release();
        pRootFolder->Release();
        CoUninitialize();
        return;
    }

    ITaskDefinition *pTaskDef = NULL;
    hr = pRegisteredTask->get_Definition(&pTaskDef);
    if (SUCCEEDED(hr))
    {
        IRegistrationInfo *pRegInfo = NULL;
        hr = pTaskDef->get_RegistrationInfo(&pRegInfo);
        if (SUCCEEDED(hr))
        {
            BSTR docs = NULL;
            hr = pRegInfo->get_Documentation(&docs);
            if (SUCCEEDED(hr))
            {
                text += QString::fromWCharArray(docs) + "\n\n";
                SysFreeString(docs);
            }
            else
            {
                qDebug() << "Cannot get the registered task docs";
            }
            pRegInfo->Release();
        }
        else
        {
            qDebug() << "Cannot get the registered task registration info";
        }
        pTaskDef->Release();
    }

    ui->taskInfoTextEdit->setText(text);
    pRootFolder->Release();
    pRegisteredTask->Release();
    pService->Release();
    CoUninitialize();
}


void MainWindow::on_runTaskButton_clicked()
{
    if (ui->taskListWidget->currentItem() == nullptr) {
        QMessageBox::critical(this, "Ошибка", "Выберите задачу");
        return;
    }

    QString name = ui->taskListWidget->currentItem()->text();
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed";
        return;
    }

    auto wname = name.toStdWString();
    LPCWSTR wszTaskName = wname.c_str();
    ITaskService *pService = NULL;
    hr = CoCreateInstance( CLSID_TaskScheduler,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ITaskService,
                          (void**)&pService );
    if (FAILED(hr))
    {
        qDebug() << "Failed to create an instance of ITaskService";
        CoUninitialize();
        return;
    }

    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() << "ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return;
    }

    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get Root Folder pointer";
        pService->Release();
        CoUninitialize();
        return;
    }


    IRegisteredTask *pRegTask = NULL;
    hr = pRootFolder->GetTask(_bstr_t( wszTaskName), &pRegTask);
    if (FAILED(hr))
    {
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    VARIANT_BOOL enabled;
    hr = pRegTask->get_Enabled(&enabled);
    if (FAILED(hr))
    {
        pRegTask->Release();
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    if (enabled == VARIANT_FALSE) {
        QMessageBox::information(this, "Ошибка", "Задача выключена");
        pRegTask->Release();
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    hr = pRegTask->Run(_variant_t(), 0);
    if (FAILED(hr))
    {
        pRegTask->Release();
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    pRegTask->Release();
    pRootFolder->Release();
    pService->Release();
    CoUninitialize();
    return;
}


void MainWindow::on_deleteTaskButton_clicked()
{
    if (ui->taskListWidget->currentItem() == nullptr) {
        QMessageBox::critical(this, "Ошибка", "Выберите задачу");
        return;
    }

    QString name = ui->taskListWidget->currentItem()->text();
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed";
        return;
    }

    auto wname = name.toStdWString();
    LPCWSTR wszTaskName = wname.c_str();
    ITaskService *pService = NULL;
    hr = CoCreateInstance( CLSID_TaskScheduler,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ITaskService,
                          (void**)&pService );
    if (FAILED(hr))
    {
        qDebug() << "Failed to create an instance of ITaskService";
        CoUninitialize();
        return;
    }

    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() << "ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return;
    }

    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get Root Folder pointer";
        pService->Release();
        CoUninitialize();
        return;
    }

    hr = pRootFolder->DeleteTask( _bstr_t( wszTaskName), 0  );
    if( FAILED(hr) )
    {
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
        return;
    }
    pRootFolder->Release();
    pService->Release();
    CoUninitialize();
    showTasks();
}


void MainWindow::on_stopTaskButton_clicked()
{
    if (ui->taskListWidget->currentItem() == nullptr) {
        QMessageBox::critical(this, "Ошибка", "Выберите задачу");
        return;
    }

    QString name = ui->taskListWidget->currentItem()->text();
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed";
        return;
    }

    auto wname = name.toStdWString();
    LPCWSTR wszTaskName = wname.c_str();
    ITaskService *pService = NULL;
    hr = CoCreateInstance( CLSID_TaskScheduler,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ITaskService,
                          (void**)&pService );
    if (FAILED(hr))
    {
        qDebug() << "Failed to create an instance of ITaskService";
        CoUninitialize();
        return;
    }

    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() << "ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return;
    }

    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get Root Folder pointer";
        pService->Release();
        CoUninitialize();
        return;
    }


    IRegisteredTask *pRegTask = NULL;
    hr = pRootFolder->GetTask(_bstr_t( wszTaskName), &pRegTask);
    if (FAILED(hr))
    {
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    VARIANT_BOOL enabled;
    hr = pRegTask->get_Enabled(&enabled);
    if (FAILED(hr))
    {
        pRegTask->Release();
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    if (enabled == VARIANT_FALSE) {
        QMessageBox::information(this, "Ошибка", "Задача выключена");
        pRegTask->Release();
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    hr = pRegTask->Stop(_variant_t());
    if (FAILED(hr))
    {
        pRegTask->Release();
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    pRegTask->Release();
    pRootFolder->Release();
    pService->Release();
    CoUninitialize();
    QMessageBox::information(this, "Успех", "Задача остановлена");
    return;
}


void MainWindow::on_disableTaskButton_clicked()
{
    if (ui->taskListWidget->currentItem() == nullptr) {
        QMessageBox::critical(this, "Ошибка", "Выберите задачу");
        return;
    }

    QString name = ui->taskListWidget->currentItem()->text();
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed";
        return;
    }

    auto wname = name.toStdWString();
    LPCWSTR wszTaskName = wname.c_str();
    ITaskService *pService = NULL;
    hr = CoCreateInstance( CLSID_TaskScheduler,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ITaskService,
                          (void**)&pService );
    if (FAILED(hr))
    {
        qDebug() << "Failed to create an instance of ITaskService";
        CoUninitialize();
        return;
    }

    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() << "ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return;
    }

    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get Root Folder pointer";
        pService->Release();
        CoUninitialize();
        return;
    }


    IRegisteredTask *pRegTask = NULL;
    hr = pRootFolder->GetTask(_bstr_t( wszTaskName), &pRegTask);
    if (FAILED(hr))
    {
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    hr = pRegTask->put_Enabled(VARIANT_FALSE);
    if (FAILED(hr))
    {
        pRegTask->Release();
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    pRegTask->Release();
    pRootFolder->Release();
    pService->Release();
    CoUninitialize();
    QMessageBox::information(this, "Ошибка", "Задача выключена");
    return;
}


void MainWindow::on_enableTaskButton_clicked()
{
    if (ui->taskListWidget->currentItem() == nullptr) {
        QMessageBox::critical(this, "Ошибка", "Выберите задачу");
        return;
    }

    QString name = ui->taskListWidget->currentItem()->text();
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed";
        return;
    }

    auto wname = name.toStdWString();
    LPCWSTR wszTaskName = wname.c_str();
    ITaskService *pService = NULL;
    hr = CoCreateInstance( CLSID_TaskScheduler,
                          NULL,
                          CLSCTX_INPROC_SERVER,
                          IID_ITaskService,
                          (void**)&pService );
    if (FAILED(hr))
    {
        qDebug() << "Failed to create an instance of ITaskService";
        CoUninitialize();
        return;
    }

    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() << "ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return;
    }

    ITaskFolder *pRootFolder = NULL;
    hr = pService->GetFolder( _bstr_t( L"\\") , &pRootFolder );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get Root Folder pointer";
        pService->Release();
        CoUninitialize();
        return;
    }


    IRegisteredTask *pRegTask = NULL;
    hr = pRootFolder->GetTask(_bstr_t( wszTaskName), &pRegTask);
    if (FAILED(hr))
    {
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    hr = pRegTask->put_Enabled(VARIANT_TRUE);
    if (FAILED(hr))
    {
        pRegTask->Release();
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return;
    }

    pRegTask->Release();
    pRootFolder->Release();
    pService->Release();
    CoUninitialize();

    QMessageBox::information(this, "Ошибка", "Задача включена");
    return;
}

