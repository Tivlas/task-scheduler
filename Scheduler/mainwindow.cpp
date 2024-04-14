#define _WIN32_DCOM
#include "mainwindow.h"
#include "./ui_mainwindow.h"
#include <windows.h>
#include <taskschd.h>
#include <comdef.h>



MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow), taskCreationPages(new pages())
{
    ui->setupUi(this);
    this->showTasks();
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

void MainWindow::showTasks() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed " << hr;
        return;
    }

    hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);

    if( FAILED(hr) )
    {
        qDebug() << "\nCoInitializeSecurity failed";
        CoUninitialize();
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

    if( numTasks == 0 )
    {
        qDebug() <<"\nNo Tasks are currently running";
        pTaskCollection->Release();
        CoUninitialize();
        return;
    }

    qDebug() << "\nNumber of Tasks:";

    //TASK_STATE taskState;

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
                QString name = QString::fromWCharArray(taskName);
                qDebug() << "\nTask Name " << name;
                ui->taskListWidget->addItem(name);
                SysFreeString(taskName);

                // hr = pRegisteredTask->get_State(&taskState);
                // if (SUCCEEDED (hr) )
                //     printf("\n\tState: %d", taskState);
                // else
                //     printf("\n\tCannot get the registered task state: %x", hr);
            }
            else
            {
                qDebug() << "\nCannot get the registered task name";
            }
            pRegisteredTask->Release();
        }
        else
        {
            qDebug() << "\nCannot get the registered task item at index";
        }
    }

    pTaskCollection->Release();
    CoUninitialize();
    return;

    // OLD below
    // ITaskService *pService = NULL;
    // CoInitialize(NULL);
    // HRESULT hr = CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService);
    // if (FAILED(hr)) {
    //     qDebug() << "Fail 1";
    //     return;
    // }


    // // Connect to the local computer task folder
    // ITaskFolder *pTaskFolder = NULL;
    // hr = pService->GetFolder(_bstr_t(L"\\"), &pTaskFolder);
    // if (FAILED(hr))
    // {
    //     pService->Release();
    //     qDebug() << "Fail 2";
    //     return;
    // }

    // // Enumerate through all the tasks in the task folder
    // IRegisteredTaskCollection *pTasks = NULL;
    // hr = pTaskFolder->GetTasks(NULL, &pTasks);
    // if (FAILED(hr))
    // {
    //     pService->Release();
    //     pTaskFolder->Release();
    //     qDebug() << "Fail 3";
    //     return;
    // }

    // long count;
    // pTasks->get_Count(&count);
    // for (long i = 0; i < count; i++)
    // {
    //     IRegisteredTask *pRegisteredTask = NULL;
    //     hr = pTasks->get_Item(_variant_t(i + 1), &pRegisteredTask);
    //     if (FAILED(hr)) {
    //         qDebug() << "Fail 4";
    //         continue;
    //     }

    //     // Get the task name
    //     BSTR taskName = NULL;
    //     hr = pRegisteredTask->get_Name(&taskName);
    //     if (SUCCEEDED(hr))
    //     {
    //         // Add the task name to the list widget
    //         QString taskNameStr = QString::fromWCharArray(taskName);
    //         ui->taskListWidget->addItem(taskNameStr);
    //         SysFreeString(taskName);
    //     } else {
    //         qDebug() << "Fail 5";
    //     }

    //     pRegisteredTask->Release();
    // }

    // pTasks->Release();
    // pTaskFolder->Release();
    // pService->Release();
    // CoUninitialize();
}
