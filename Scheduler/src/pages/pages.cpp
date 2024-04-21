#include "pages.h"
#include "ui_pages.h"

pages::pages(QWidget *parent)
    : QDialog(parent)
    , ui(new Ui::pages)
{
    ui->setupUi(this);
    ui->addTaskButton->hide();
    ui->cancelButton->hide();
    ui->onceRadioButton->click();
    ui->createPages->setCurrentIndex(0);
    ui->startTimeEdit->setMinimumDateTime(QDateTime::currentDateTime());
}

pages::~pages()
{
    delete ui;
}

void pages::warningMsgBox(QString title, QString text) {
    QMessageBox::warning(this, title, text);
}

void pages::errorMsgBox(QString title, QString text) {
    QMessageBox::critical(this, title, text);
}

void pages::okMsgBox(QString title, QString text) {
    QMessageBox::information(this, title, text);
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
    pages::ErrCode err = pages::Ok;
    if (ui->onceRadioButton->isChecked()) {
        err = addSpecificTimeTask();
    } else if (ui->dailyRadioButton->isChecked()) {
        err = addDailyTask();
    } else if (ui->weeklyRadioButton->isChecked()) {
        err = addWeeklyTask();
    }
    if (err == pages::Ok) {
        this->done(DONE_CODE);
        ui->createPages->setCurrentIndex(0);
    }
}

pages::ErrCode pages::addDailyTask() {
    QString name = getTaskName();
    QString description = getTaskDescription();
    QString actionPath = getActionPath();
    QString actionArgs = getActionArgs();
    QDateTime startDateTime = getStartDateTime();

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed";
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Create a name for the task.
    auto wname = name.toStdWString();
    LPCWSTR wszTaskName = wname.c_str();

    std::wstring wstrExecutablePath = L"";
    auto wactionPath = actionPath.toStdWString();
    wactionPath += L" " + actionArgs.toStdWString();
    wstrExecutablePath += wactionPath.c_str();

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
        return pages::Err;
    }

    //  Connect to the task service.
    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() << "ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return pages::Err;
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
        return pages::Err;
    }

    // If the same task exists, remove it.
    IRegisteredTask *pExistsTask = NULL;
    hr = pRootFolder->GetTask(_bstr_t( wszTaskName), &pExistsTask);
    if (SUCCEEDED(hr))
    {
        pExistsTask->Release();
        QMessageBox::StandardButton reply = QMessageBox::question(this, "Внимание", "Задача с таким именем уже существует, хотите ее удалить?");
        if (reply == QMessageBox::No)
        {
            pService->Release();
            CoUninitialize();
            return pages::Err;
        }
    }

    pRootFolder->DeleteTask( _bstr_t( wszTaskName), 0  );

    //  Create the task builder object to create the task.
    ITaskDefinition *pTask = NULL;
    hr = pService->NewTask( 0, &pTask );

    pService->Release();  // COM clean up.  Pointer is no longer used.
    if (FAILED(hr))
    {
        qDebug() << "Failed to CoCreate an instance of the TaskService class";
        pRootFolder->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Get the registration info for setting the identification.
    IRegistrationInfo *pRegInfo= NULL;
    hr = pTask->get_RegistrationInfo( &pRegInfo );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get identification pointer";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    hr = pRegInfo->put_Author( _bstr_t(L"DEFAULT") );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot put author info";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    std::wstring wdescription = description.toStdWString();
    hr = pRegInfo->put_Description(_bstr_t(wdescription.c_str()) );
    pRegInfo->Release();  // COM clean up.  Pointer is no longer used.
    if( FAILED(hr) )
    {
        qDebug() << "Cannot put description info";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Get the trigger collection to insert the daily trigger.
    ITriggerCollection *pTriggerCollection = NULL;
    hr = pTask->get_Triggers( &pTriggerCollection );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get trigger collection";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  Add the daily trigger to the task.
    ITrigger *pTrigger = NULL;
    hr = pTriggerCollection->Create( TASK_TRIGGER_DAILY, &pTrigger );
    pTriggerCollection->Release();
    if( FAILED(hr) )
    {
        qDebug() << "Cannot create the trigger";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    IDailyTrigger *pDailyTrigger = NULL;
    hr = pTrigger->QueryInterface(
        IID_IDailyTrigger, (void**) &pDailyTrigger );
    pTrigger->Release();
    if( FAILED(hr) )
    {
        qDebug() << "QueryInterface call on IDailyTrigger failed";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    std::wstring triggerId = L"daily_" + wname + startDateTime.toString().toStdWString();
    hr = pDailyTrigger->put_Id( _bstr_t( triggerId.c_str()) );
    if( FAILED(hr) )
        qDebug() << "Cannot put trigger ID";

    //  Set the task to start daily at a certain time. The time
    //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
    //  For example, the start boundary below
    //  is January 1st 2005 at 12:05
    std::wstring dateTimeFormatted = startDateTime.toString("yyyy-MM-dd'T'hh:mm:ss").toStdWString();
    hr = pDailyTrigger->put_StartBoundary( _bstr_t(dateTimeFormatted.c_str()) );
    if( FAILED(hr) )
        qDebug() << "Cannot put start boundary";

    // TODO: let user choose this.
    //  Set the time when the trigger is deactivated.
    // hr = pDailyTrigger->put_EndBoundary( _bstr_t(L"2007-05-02T12:05:00") );
    // if( FAILED(hr) )
    //     qDebug() << "\nCannot put the end boundary";

    //  Define the interval for the daily trigger. An interval of 2 produces an
    //  every other day schedule
    bool ok;
    auto daysInterval = ui->everyNthDayLineEdit->text().toInt(&ok);
    if (!ok || daysInterval <= 0) {
        errorMsgBox("Ошибка", "Вы не ввели/ввели некорректную частоту выполнения задачи.");
        return pages::Err;
    }
    hr = pDailyTrigger->put_DaysInterval( static_cast<short>(daysInterval) );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot put days interval";
        pRootFolder->Release();
        pDailyTrigger->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    // TODO: let user choose this.
    // Add a repetition to the trigger so that it repeats
    // five times.
    // IRepetitionPattern *pRepetitionPattern = NULL;
    // hr = pDailyTrigger->get_Repetition( &pRepetitionPattern );
    // pDailyTrigger->Release();
    // if( FAILED(hr) )
    // {
    //     qDebug() << "Cannot get repetition pattern";
    //     pRootFolder->Release();
    //     pTask->Release();
    //     CoUninitialize();
    //     return pages::Err;
    // }

    // hr = pRepetitionPattern->put_Duration( _bstr_t(L"PT4M"));
    // if( FAILED(hr) )
    // {
    //     qDebug() << "Cannot put repetition duration";
    //     pRootFolder->Release();
    //     pRepetitionPattern->Release();
    //     pTask->Release();
    //     CoUninitialize();
    //     return pages::Err;
    // }

    // hr = pRepetitionPattern->put_Interval( _bstr_t(L"PT1M"));
    // pRepetitionPattern->Release();
    // if( FAILED(hr) )
    // {
    //     qDebug() << "Cannot put repetition interval";
    //     pRootFolder->Release();
    //     pTask->Release();
    //     CoUninitialize();
    //     return pages::Err;
    // }

    //  ------------------------------------------------------
    //  Add an action to the task. This task will execute notepad.exe.
    IActionCollection *pActionCollection = NULL;

    //  Get the task action collection pointer.
    hr = pTask->get_Actions( &pActionCollection );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get task collection pointer";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  Create the action, specifying that it is an executable action.
    IAction *pAction = NULL;
    hr = pActionCollection->Create( TASK_ACTION_EXEC, &pAction );
    pActionCollection->Release();
    if( FAILED(hr) )
    {
        qDebug() << "Cannot create action";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    IExecAction *pExecAction = NULL;
    hr = pAction->QueryInterface(
        IID_IExecAction, (void**) &pExecAction );
    pAction->Release();
    if( FAILED(hr) )
    {
        qDebug() << "QueryInterface call failed for IExecAction";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  Set the path of the executable to notepad.exe.
    hr = pExecAction->put_Path( _bstr_t( wstrExecutablePath.c_str() ) );
    pExecAction->Release();
    if( FAILED(hr) )
    {
        qDebug() << "Cannot put the executable path";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Save the task in the root folder.
    IRegisteredTask *pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t( wszTaskName ),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(),
        _variant_t(),
        TASK_LOGON_NONE,
        _variant_t(L""),
        &pRegisteredTask);
    if( FAILED(hr) )
    {
        qDebug() << "Error saving the Task";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    qDebug() << "Success! Task successfully registered.";

    //  Clean up
    pRootFolder->Release();
    pTask->Release();
    pRegisteredTask->Release();
    CoUninitialize();
    return pages::Ok;
}

pages::ErrCode pages::addSpecificTimeTask() {
    QString name = getTaskName();
    QString description = getTaskDescription();
    QString actionPath = getActionPath();
    QString actionArgs = getActionArgs();
    QDateTime startDateTime = getStartDateTime();

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if( FAILED(hr) )
    {
        qDebug() << "CoInitializeEx failed";
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Create a name for the task.
    auto wname = name.toStdWString();
    LPCWSTR wszTaskName = wname.c_str();

    std::wstring wstrExecutablePath = L"";
    auto wactionPath = actionPath.toStdWString();
    wactionPath += L" " + actionArgs.toStdWString();
    wstrExecutablePath += wactionPath.c_str();

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
        return pages::Err;
    }

    //  Connect to the task service.
    hr = pService->Connect(_variant_t(), _variant_t(),
                           _variant_t(), _variant_t());
    if( FAILED(hr) )
    {
        qDebug() << "ITaskService::Connect failed";
        pService->Release();
        CoUninitialize();
        return pages::Err;
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
        return pages::Err;
    }

    // If the same task exists, remove it.
    IRegisteredTask *pExistsTask = NULL;
    hr = pRootFolder->GetTask(_bstr_t( wszTaskName), &pExistsTask);
    if (SUCCEEDED(hr))
    {
        pExistsTask->Release();
        QMessageBox::StandardButton reply = QMessageBox::question(this, "Внимание", "Задача с таким именем уже существует, хотите ее удалить?");
        if (reply == QMessageBox::No)
        {
            pService->Release();
            CoUninitialize();
            return pages::Err;
        }
    }

    pRootFolder->DeleteTask( _bstr_t( wszTaskName), 0  );

    //  Create the task builder object to create the task.
    ITaskDefinition *pTask = NULL;
    hr = pService->NewTask( 0, &pTask );

    pService->Release();  // COM clean up.  Pointer is no longer used.
    if (FAILED(hr))
    {
        qDebug() << "Failed to CoCreate an instance of the TaskService class";
        pRootFolder->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Get the registration info for setting the identification.
    IRegistrationInfo *pRegInfo= NULL;
    hr = pTask->get_RegistrationInfo( &pRegInfo );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get identification pointer";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    // hr = pRegInfo->put_Author( _bstr_t(L"DEFAULT") );
    // if( FAILED(hr) )
    // {
    //     qDebug() << "Cannot put author info";
    //     pRootFolder->Release();
    //     pTask->Release();
    //     CoUninitialize();
    //     return pages::Err;
    // }

    std::wstring wdescription = description.toStdWString();
    hr = pRegInfo->put_Description(_bstr_t(wdescription.c_str()) );
    pRegInfo->Release();  // COM clean up.  Pointer is no longer used.
    if( FAILED(hr) )
    {
        qDebug() << "Cannot put description info";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    ITaskSettings *pSettings = NULL;
    hr = pTask->get_Settings( &pSettings );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get settings pointer";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  Set setting values for the task.
    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    pSettings->Release();
    if( FAILED(hr) )
    {
        qDebug() << "Cannot put setting information";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Get the trigger collection to insert the daily trigger.
    ITriggerCollection *pTriggerCollection = NULL;
    hr = pTask->get_Triggers( &pTriggerCollection );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get trigger collection";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  Add the time trigger to the task.
    ITrigger *pTrigger = NULL;
    hr = pTriggerCollection->Create( TASK_TRIGGER_TIME, &pTrigger );
    pTriggerCollection->Release();
    if( FAILED(hr) )
    {
        qDebug() << "Cannot create the trigger";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    ITimeTrigger *pTimeTrigger = NULL;
    hr = pTrigger->QueryInterface(
        IID_ITimeTrigger, (void**) &pTimeTrigger );
    pTrigger->Release();
    if( FAILED(hr) )
    {
        qDebug() << "QueryInterface call on ITimeTrigger failed";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    std::wstring triggerId = L"time_" + wname + startDateTime.toString().toStdWString();
    hr = pTimeTrigger->put_Id( _bstr_t( triggerId.c_str()) );
    if( FAILED(hr) )
        qDebug() << "Cannot put trigger ID";

    //  Set the task to start daily at a certain time. The time
    //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
    //  For example, the start boundary below
    //  is January 1st 2005 at 12:05
    std::wstring dateTimeFormatted = startDateTime.toString("yyyy-MM-dd'T'hh:mm:ss").toStdWString();
    hr = pTimeTrigger->put_StartBoundary( _bstr_t(dateTimeFormatted.c_str()) );
    if( FAILED(hr) )
        qDebug() << "Cannot put start boundary";

    // TODO: let user set end boundary.
    bool ok;
    auto daysInterval = ui->everyNthDayLineEdit->text().toInt(&ok);
    if (!ok || daysInterval <= 0) {
        errorMsgBox("Ошибка", "Вы не ввели/ввели некорректную частоту выполнения задачи.");
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Add an action to the task. This task will execute notepad.exe.
    IActionCollection *pActionCollection = NULL;

    //  Get the task action collection pointer.
    hr = pTask->get_Actions( &pActionCollection );
    if( FAILED(hr) )
    {
        qDebug() << "Cannot get task collection pointer";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  Create the action, specifying that it is an executable action.
    IAction *pAction = NULL;
    hr = pActionCollection->Create( TASK_ACTION_EXEC, &pAction );
    pActionCollection->Release();
    if( FAILED(hr) )
    {
        qDebug() << "Cannot create action";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    IExecAction *pExecAction = NULL;
    hr = pAction->QueryInterface(
        IID_IExecAction, (void**) &pExecAction );
    pAction->Release();
    if( FAILED(hr) )
    {
        qDebug() << "QueryInterface call failed for IExecAction";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  Set the path of the executable to notepad.exe.
    hr = pExecAction->put_Path( _bstr_t( wstrExecutablePath.c_str() ) );
    pExecAction->Release();
    if( FAILED(hr) )
    {
        qDebug() << "Cannot put the executable path";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    //  ------------------------------------------------------
    //  Save the task in the root folder.
    IRegisteredTask *pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t( wszTaskName ),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(),
        _variant_t(),
        TASK_LOGON_NONE,
        _variant_t(L""),
        &pRegisteredTask);
    if( FAILED(hr) )
    {
        qDebug() << "Error saving the Task";
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return pages::Err;
    }

    qDebug() << "Success! Task successfully registered.";

    //  Clean up
    pRootFolder->Release();
    pTask->Release();
    pRegisteredTask->Release();
    CoUninitialize();
    return pages::Ok;
}

pages::ErrCode pages::addWeeklyTask() {
    return pages::Ok;
}

QString pages::getTaskName() {
    return ui->nameLineEdit->text();
}

QString pages::getTaskDescription() {
    return ui->taskDescriptionTextEdit->toPlainText();
}

QString pages::getActionPath() {
    return ui->pathToActionLineEdit->text();
}

QString pages::getActionArgs() {
    return ui->actionArgsLineEdit->text();
}

QDateTime pages::getStartDateTime() {
    return ui->startTimeEdit->dateTime();
}

void pages::on_cancelButton_clicked()
{
    ui->createPages->setCurrentIndex(0);
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

