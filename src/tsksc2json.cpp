#include <windows.h>
#include <taskschd.h>
#include <comdef.h>
#include <atlbase.h>
#include <iostream>
#include <fstream>
#include <numeric>
#include <string>
#include <algorithm>
#include <vector>
#include <format>
#include <regex>

#pragma comment( lib, "taskschd" )
//
static std::wstring join(std::vector<std::wstring>& strings, const std::wstring& delim)
{
    return std::accumulate(strings.begin() + 1, strings.end(), strings[0],
        [&delim](const std::wstring& x, const std::wstring& y) {
            return x.empty() ? y : x + delim + y;
        });
}
static std::wstring bstr_to_string(BSTR bstr)
{
    _bstr_t b(bstr);
    return std::wstring((const wchar_t*)b);
}
static std::wstring time_to_string(const DATE& date)
{
    SYSTEMTIME st;
    VariantTimeToSystemTime(date, &st);
    wchar_t buffer[32];
    swprintf_s(buffer, L"%04d/%02d/%02d %02d:%02d:%02d",
        st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    return buffer;
}
static std::wstring escape_string(std::wstring str)
{
    std::wstring ret = std::regex_replace(str, std::wregex(L"(\\\\|\")"), L"\\$1");
    ret = std::regex_replace(ret, std::wregex(L"(\r\n|\r|\n)"), L"\\n");
    return L"\"" + ret + L"\"";
}
static std::wstring escape_string(BSTR str)
{
    return escape_string(bstr_to_string(str));
}
static std::wstring escape_string(DATE date)
{
    return escape_string(time_to_string(date));
}
static void add_params(std::vector<std::wstring>& params, std::vector<std::wstring>& sub_params, std::wstring indent, std::wstring name = L"")
{
    if (sub_params.size() > 0)
    {
        params.push_back(std::format(L"{}{{\n {}{}\n{}}}", name.empty() ? indent : name, indent, join(sub_params, std::format(L",\n {}", indent)), indent));
    }
}
static void output_params(std::vector<std::wstring>& params, std::wstring name)
{
    if (params.size() > 0)
    {
        std::wcout << L",\n \"" << name << L"\": [\n" << join(params, L",\n") << L"\n ]";
    }
}
//
static void extract_registration_info(ITaskDefinition* taskdef)
{
    CComPtr<IRegistrationInfo> info;
    if (FAILED(taskdef->get_RegistrationInfo(&info)) || !info)
    {
        return;
    }
    if (BSTR desc; SUCCEEDED(info->get_Description(&desc)) && desc)
    {
        std::wcout << L",\n \"Description\": " << escape_string(desc);
    }
}
static void extract_actions(ITaskDefinition* taskdef)
{
    CComPtr<IActionCollection> actions;
    LONG count;
    if (FAILED(taskdef->get_Actions(&actions)) || !actions
        || FAILED(actions->get_Count(&count)) || count <= 0
       )
    {
        return;
    }
    std::vector<std::wstring> actions_params;
    for (LONG i = 1; i <= count; ++i)
    {
        CComPtr<IAction> action;
        TASK_ACTION_TYPE type;
        CComPtr<IExecAction> exec;
        if (FAILED(actions->get_Item(i, &action)) || !action
            || FAILED(action->get_Type(&type)) || type != TASK_ACTION_EXEC
            || FAILED(action->QueryInterface(IID_IExecAction, (void**)&exec)) || !exec
           )
        {
            continue;
        }
        std::vector<std::wstring> action_params;
        if (BSTR cmd; SUCCEEDED(exec->get_Path(&cmd)) && cmd)
        {
            action_params.push_back(std::format(L"\"Command\": {}", escape_string(cmd)));
        }
        if (BSTR args; SUCCEEDED(exec->get_Arguments(&args)) && args)
        {
            action_params.push_back(std::format(L"\"Arguments\": {}", escape_string(args)));
        }
        if (BSTR dir; SUCCEEDED(exec->get_WorkingDirectory(&dir)) && dir)
        {
            action_params.push_back(std::format(L"\"WorkingDirectory\": {}", escape_string(dir)));
        }
        add_params(actions_params, action_params, L"  ");
    }
    output_params(actions_params, L"Actions");
}
//
static void extract_triggers(ITaskDefinition* taskdef)
{
    CComPtr<ITriggerCollection> triggers;
    LONG count;
    if (FAILED(taskdef->get_Triggers(&triggers)) || !triggers
        || FAILED(triggers->get_Count(&count)) || count <= 0
       )
    {
        return;
    }
    std::vector<std::wstring> triggers_params;
    for (LONG i = 1; i <= count; ++i)
    {
        CComPtr<ITrigger> trigger;
        if (FAILED(triggers->get_Item(i, &trigger)) || !trigger)
        {
            continue;
        }
        std::vector<std::wstring> trigger_params;
        if (BSTR start; SUCCEEDED(trigger->get_StartBoundary(&start)) && start)
        {
            trigger_params.push_back(std::format(L"\"StartBoundary\": {}", escape_string(start)));
        }
        if (VARIANT_BOOL enabled; SUCCEEDED(trigger->get_Enabled(&enabled)))
        {
            trigger_params.push_back(std::format(L"\"Enabled\": {}", enabled == VARIANT_TRUE));
        }
        if (TASK_TRIGGER_TYPE2 ttype; SUCCEEDED(trigger->get_Type(&ttype)))
        {
            switch(ttype)
            {
            case TASK_TRIGGER_DAILY: // Daily
                if (CComPtr<IDailyTrigger> daily; SUCCEEDED(trigger->QueryInterface(IID_IDailyTrigger, (void**)&daily)))
                {
                    if (SHORT interval; SUCCEEDED(daily->get_DaysInterval(&interval)))
                    {
                        trigger_params.push_back(std::format(L"\"ScheduleByDay\": {{ \"DaysInterval\":{} }}", interval));
                    }
                }
                break;
            case TASK_TRIGGER_WEEKLY: // Weekly
                if (CComPtr<IWeeklyTrigger> weekly; SUCCEEDED(trigger->QueryInterface(IID_IWeeklyTrigger, (void**)&weekly)))
                {
                    std::vector<std::wstring> weekly_params;
                    if (SHORT interval; SUCCEEDED(weekly->get_WeeksInterval(&interval)))
                    {
                        weekly_params.push_back(std::format(L"\"WeeksInterval\": {}", interval));
                    }
                    if (short daysOfWeek; SUCCEEDED(weekly->get_DaysOfWeek(&daysOfWeek)))
                    {
                        weekly_params.push_back(std::format(L"\"DaysOfWeek\": {}", daysOfWeek));
                    }
                    add_params(trigger_params, weekly_params, L"   ", L"\"ScheduleByWeek\": ");
                }
                break;
            case TASK_TRIGGER_MONTHLY: // Monthly
                if (CComPtr<IMonthlyTrigger> monthly; SUCCEEDED(trigger->QueryInterface(IID_IMonthlyTrigger, (void**)&monthly)))
                {
                    std::vector<std::wstring> monthly_params;
                    if (long days; SUCCEEDED(monthly->get_DaysOfMonth(&days)))
                    {
                        monthly_params.push_back(std::format(L"\"Days\": {}", days));
                    }
                    if (short months; SUCCEEDED(monthly->get_MonthsOfYear(&months)))
                    {
                        monthly_params.push_back(std::format(L"\"Months\": {}", months));
                    }
                    if (VARIANT_BOOL lastDay = VARIANT_FALSE; SUCCEEDED(monthly->get_RunOnLastDayOfMonth(&lastDay)) && lastDay == VARIANT_TRUE)
                    {
                        monthly_params.push_back(L"\"RunOnLastDayOfMonth\": true");
                    }
                    add_params(trigger_params, monthly_params, L"   ", L"\"ScheduleByMonth\": ");
                }
                break;
            }
        }
        // Repetition
        if (CComPtr<IRepetitionPattern> rep; SUCCEEDED(trigger->get_Repetition(&rep)) && rep)
        {
            std::vector<std::wstring> repetition_params;
            if (BSTR interval; SUCCEEDED(rep->get_Interval(&interval)) && interval)
            {
                repetition_params.push_back(std::format(L"\"Interval\": {}", escape_string(interval)));
            }
            if (BSTR duration; SUCCEEDED(rep->get_Duration(&duration)) && duration)
            {
                repetition_params.push_back(std::format(L"\"Duration\": {}", escape_string(duration)));
            }
            if (VARIANT_BOOL stop; SUCCEEDED(rep->get_StopAtDurationEnd(&stop)) && (stop == VARIANT_TRUE))
            {
                repetition_params.push_back(L"\"StopAtDurationEnd\": true");
            }
            add_params(trigger_params, repetition_params, L"   ", L"\"Repetition\": ");
        }
        add_params(triggers_params, trigger_params, L"  ");
    }
    output_params(triggers_params, L"Triggers");
}
//
static void task_info(CComPtr<IRegisteredTask> &task, const std::wstring& path)
{
    BSTR name;
    TASK_STATE state;
    DATE lastRun, nextRun;
    LONG result;
    if (SUCCEEDED(task->get_Name(&name))
        && SUCCEEDED(task->get_State(&state))
        && SUCCEEDED(task->get_LastRunTime(&lastRun))
        && SUCCEEDED(task->get_NextRunTime(&nextRun))
        && SUCCEEDED(task->get_LastTaskResult(&result))
        )
    {
        std::wcout
            << L"{\n"
            << L" \"TaskPath\": " << escape_string(path) << L",\n"
            << L" \"TaskName\": " << escape_string(name) << L",\n"
            << L" \"State\": " << static_cast<int>(state) << L",\n"
            << L" \"LastRunTime\": " << escape_string(lastRun) << L",\n"
            << L" \"NextRunTime\": " << escape_string(nextRun) << L",\n"
            << L" \"LastResult\": " << result
            ;
        if (ITaskDefinition* taskdef; SUCCEEDED(task->get_Definition(&taskdef)) && taskdef)
        {
            extract_registration_info(taskdef);
            extract_actions(taskdef);
            extract_triggers(taskdef);
        }
        std::wcout << L"\n}";
    }
    else
    {
        std::wcout << L"{  \"TaskPath\": " << escape_string(path) << L"}\n"; // Error
    }
}
static int enumerate_tasks_in_folder(ITaskFolder* folder, const std::wstring& path, const int in_element_count = 0)
{
    CComPtr<IRegisteredTaskCollection> tasks;
    if (FAILED(folder->GetTasks(TASK_ENUM_HIDDEN, &tasks)) || !tasks)
    {
        return 0;
    }
    int element_count = in_element_count;
    std::wstring del_str = element_count == 0 ? L"[" : L",";
    if (LONG count; SUCCEEDED(tasks->get_Count(&count)))
    {
        for (LONG i = 1; i <= count; ++i)
        {
            if (CComPtr<IRegisteredTask> task; SUCCEEDED(tasks->get_Item(_variant_t(i), &task)) && task)
            {
                std::wcout << del_str;
                del_str = L",";
                task_info(task, path);
                ++element_count;
            }
        }
    }
    // 再帰: サブフォルダ探索
    if (CComPtr<ITaskFolderCollection> subfolders; SUCCEEDED(folder->GetFolders(0, &subfolders)) && subfolders)
    {
        if (LONG subCount; SUCCEEDED(subfolders->get_Count(&subCount)))
        {
            for (LONG i = 1; i <= subCount; ++i)
            {
                if (CComPtr<ITaskFolder> subfolder; SUCCEEDED(subfolders->get_Item(_variant_t(i), &subfolder)) && subfolder)
                {
                    if (BSTR subPath; SUCCEEDED(subfolder->get_Path(&subPath)) && subPath)
                    {
                        element_count += enumerate_tasks_in_folder(subfolder, bstr_to_string(subPath), element_count);
                    }
                }
            }
        }
    }
    if (element_count == 0)
    {
        std::wcout << L"[]\n";
    }
    else if (in_element_count == 0)
    {
        std::wcout << L"]\n";
    }
    return element_count;
}
//
int wmain(int argc, wchar_t* argv[])
{
    std::locale::global(std::locale(""));
    std::wstring rootPath = L"\\";  // デフォルトルート
    if (argc > 1)
    {
        rootPath = argv[1];
    }
    if (FAILED(CoInitializeEx(nullptr, COINIT_MULTITHREADED)))
    {
        return EXIT_FAILURE;
    }
    if (FAILED(CoInitializeSecurity(nullptr, -1, nullptr, nullptr,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, 0, nullptr)))
    {
        CoUninitialize();
        return EXIT_FAILURE;
    }
    if (CComPtr<ITaskService> service;
        SUCCEEDED(service.CoCreateInstance(CLSID_TaskScheduler))
        && SUCCEEDED(service->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t()))
        )
    {
        //std::wofstream writing_file(L"test.json", std::ios::out | std::ios::binary);
        //auto strbuf = std::wcout.rdbuf(writing_file.rdbuf());
        if (rootPath.at(rootPath.size() - 1) == L'\\')
        {
            if (rootPath.size() > 1)
            {
                rootPath.erase(rootPath.size() - 1, 1); // 最後の1文字を削除
            }
            if (CComPtr<ITaskFolder> root; SUCCEEDED(service->GetFolder(_bstr_t(rootPath.c_str()), &root)) && root)
            {
                enumerate_tasks_in_folder(root, rootPath);
            }
        }
        else if (std::wsmatch m; std::regex_match(rootPath, m, std::wregex(L"^(.*)\\\\(.*)$")))
        {
            std::wstring path = m[1].str();
            std::wstring name = m[2].str();
            if (CComPtr<ITaskFolder> root; SUCCEEDED(service->GetFolder(_bstr_t(path.c_str()), &root)) && root)
            {
                if (CComPtr<IRegisteredTask> task; SUCCEEDED(root->GetTask(_bstr_t(name.c_str()), &task)) && task)
                {
                    std::wcout << L"[";
                    task_info(task, path);
                    std::wcout << L"]\n";
                }
            }
        }
        //std::wcout.rdbuf(strbuf);
    }
    CoUninitialize();
}