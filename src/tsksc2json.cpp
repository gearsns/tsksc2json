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
static std::string join(std::vector<std::string>& strings, const std::string& delim)
{
    return std::accumulate(strings.begin() + 1, strings.end(), strings[0],
        [&delim](const std::string& x, const std::string& y) {
            return x.empty() ? y : x + delim + y;
        });
}
static std::string bstr_to_utf8(BSTR bstr)
{
    _bstr_t b(bstr);
    return std::string((const char*)b);
}
/*
static std::string bstr_to_utf8(const std::wstring& wstr)
{
    if (wstr.empty())
    {
        return "";
    }
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string result(size_needed - 1, 0); // -1 to remove null terminator
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, result.data(), size_needed, nullptr, nullptr);
    return result;
}
*/
static std::string time_to_string(const DATE& date)
{
    SYSTEMTIME st;
    VariantTimeToSystemTime(date, &st);
    char buffer[32];
    sprintf_s(buffer, "%04d/%02d/%02d %02d:%02d:%02d",
        st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    return buffer;
}
static std::string escape_string(std::string str)
{
    std::string ret = std::regex_replace(str, std::regex("(\\\\|\")"), "\\$1");
    ret = std::regex_replace(ret, std::regex("(\r\n|\r|\n)"), "\\n");
    return "\"" + ret + "\"";
}
static std::string escape_string(BSTR str)
{
    return escape_string(bstr_to_utf8(str));
}
static std::string escape_string(DATE date)
{
    return escape_string(time_to_string(date));
}
static void add_params(std::vector<std::string>& params, std::vector<std::string>& sub_params, std::string indent, std::string name = "")
{
    if (sub_params.size() > 0)
    {
        params.push_back(std::format("{}{{\n {}{}\n{}}}", name.empty() ? indent : name, indent, join(sub_params, std::format(",\n {}", indent)), indent));
    }
}
static void output_params(std::vector<std::string>& params, std::string name)
{
    if (params.size() > 0)
    {
        std::cout << ",\n \"" << name << "\": [\n" << join(params, ",\n") << "\n ]";
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
        std::cout << ",\n \"Description\": " << escape_string(desc);
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
    std::vector<std::string> actions_params;
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
        std::vector<std::string> action_params;
        if (BSTR cmd; SUCCEEDED(exec->get_Path(&cmd)) && cmd)
        {
            action_params.push_back(std::format("\"Command\": {}", escape_string(cmd)));
        }
        if (BSTR args; SUCCEEDED(exec->get_Arguments(&args)) && args)
        {
            action_params.push_back(std::format("\"Arguments\": {}", escape_string(args)));
        }
        if (BSTR dir; SUCCEEDED(exec->get_WorkingDirectory(&dir)) && dir)
        {
            action_params.push_back(std::format("\"WorkingDirectory\": {}", escape_string(dir)));
        }
        add_params(actions_params, action_params, "  ");
    }
    output_params(actions_params, "Actions");
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
    std::vector<std::string> triggers_params;
    for (LONG i = 1; i <= count; ++i)
    {
        CComPtr<ITrigger> trigger;
        if (FAILED(triggers->get_Item(i, &trigger)) || !trigger)
        {
            continue;
        }
        std::vector<std::string> trigger_params;
        if (BSTR start; SUCCEEDED(trigger->get_StartBoundary(&start)) && start)
        {
            trigger_params.push_back(std::format("\"StartBoundary\": {}", escape_string(start)));
        }
        if (VARIANT_BOOL enabled; SUCCEEDED(trigger->get_Enabled(&enabled)))
        {
            trigger_params.push_back(std::format("\"Enabled\": {}", enabled == VARIANT_TRUE));
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
                        trigger_params.push_back(std::format("\"ScheduleByDay\": {{ \"DaysInterval\":{} }}", interval));
                    }
                }
                break;
            case TASK_TRIGGER_WEEKLY: // Weekly
                if (CComPtr<IWeeklyTrigger> weekly; SUCCEEDED(trigger->QueryInterface(IID_IWeeklyTrigger, (void**)&weekly)))
                {
                    std::vector<std::string> weekly_params;
                    if (SHORT interval; SUCCEEDED(weekly->get_WeeksInterval(&interval)))
                    {
                        weekly_params.push_back(std::format("\"WeeksInterval\": {}", interval));
                    }
                    if (short daysOfWeek; SUCCEEDED(weekly->get_DaysOfWeek(&daysOfWeek)))
                    {
                        weekly_params.push_back(std::format("\"DaysOfWeek\": {}", daysOfWeek));
                    }
                    add_params(trigger_params, weekly_params, "   ", "\"ScheduleByWeek\": ");
                }
                break;
            case TASK_TRIGGER_MONTHLY: // Monthly
                if (CComPtr<IMonthlyTrigger> monthly; SUCCEEDED(trigger->QueryInterface(IID_IMonthlyTrigger, (void**)&monthly)))
                {
                    std::vector<std::string> monthly_params;
                    if (long days; SUCCEEDED(monthly->get_DaysOfMonth(&days)))
                    {
                        monthly_params.push_back(std::format("\"Days\": {}", days));
                    }
                    if (short months; SUCCEEDED(monthly->get_MonthsOfYear(&months)))
                    {
                        monthly_params.push_back(std::format("\"Months\": {}", months));
                    }
                    if (VARIANT_BOOL lastDay = VARIANT_FALSE; SUCCEEDED(monthly->get_RunOnLastDayOfMonth(&lastDay)) && lastDay == VARIANT_TRUE)
                    {
                        monthly_params.push_back("\"RunOnLastDayOfMonth\": true");
                    }
                    add_params(trigger_params, monthly_params, "   ", "\"ScheduleByMonth\": ");
                }
                break;
            }
        }
        // Repetition
        if (CComPtr<IRepetitionPattern> rep; SUCCEEDED(trigger->get_Repetition(&rep)) && rep)
        {
            std::vector<std::string> repetition_params;
            if (BSTR interval; SUCCEEDED(rep->get_Interval(&interval)) && interval)
            {
                repetition_params.push_back(std::format("\"Interval\": {}", escape_string(interval)));
            }
            if (BSTR duration; SUCCEEDED(rep->get_Duration(&duration)) && duration)
            {
                repetition_params.push_back(std::format("\"Duration\": {}", escape_string(duration)));
            }
            if (VARIANT_BOOL stop; SUCCEEDED(rep->get_StopAtDurationEnd(&stop)) && (stop == VARIANT_TRUE))
            {
                repetition_params.push_back("\"StopAtDurationEnd\": true");
            }
            add_params(trigger_params, repetition_params, "   ", "\"Repetition\": ");
        }
        add_params(triggers_params, trigger_params, "  ");
    }
    output_params(triggers_params, "Triggers");
}
//
static void task_info(CComPtr<IRegisteredTask> &task, const std::string& path)
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
        std::cout
            << "{\n"
            << " \"TaskPath\": " << escape_string(path) << ",\n"
            << " \"TaskName\": " << escape_string(name) << ",\n"
            << " \"State\": " << static_cast<int>(state) << ",\n"
            << " \"LastRunTime\": " << escape_string(lastRun) << ",\n"
            << " \"NextRunTime\": " << escape_string(nextRun) << ",\n"
            << " \"LastResult\": " << result
            ;
        if (ITaskDefinition* taskdef; SUCCEEDED(task->get_Definition(&taskdef)) && taskdef)
        {
            extract_registration_info(taskdef);
            extract_actions(taskdef);
            extract_triggers(taskdef);
        }
        std::cout << "\n}";
    }
    else
    {
        std::cout << "{ " << " \"TaskPath\": " << escape_string(path) << "}\n"; // Error
    }
}
static int enumerate_tasks_in_folder(ITaskFolder* folder, const std::string& path, const int in_element_count = 0)
{
    CComPtr<IRegisteredTaskCollection> tasks;
    if (FAILED(folder->GetTasks(TASK_ENUM_HIDDEN, &tasks)) || !tasks)
    {
        return 0;
    }
    int element_count = in_element_count;
    char del_str = element_count == 0 ? '[' : ',';
    if (LONG count; SUCCEEDED(tasks->get_Count(&count)))
    {
        for (LONG i = 1; i <= count; ++i)
        {
            if (CComPtr<IRegisteredTask> task; SUCCEEDED(tasks->get_Item(_variant_t(i), &task)) && task)
            {
                std::cout << del_str;
                del_str = ',';
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
                        element_count += enumerate_tasks_in_folder(subfolder, bstr_to_utf8(subPath), element_count);
                    }
                }
            }
        }
    }
    if (element_count == 0)
    {
        std::cout << "[]\n";
    }
    else if (in_element_count == 0)
    {
        std::cout << "]\n";
    }
    return element_count;
}
//
int main(int argc, char* argv[])
{
    std::locale::global(std::locale(""));
    std::string rootPath = "\\";  // デフォルトルート
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
    UINT saveCP = GetConsoleOutputCP();
    SetConsoleOutputCP(CP_UTF8);
    if (CComPtr<ITaskService> service;
        SUCCEEDED(service.CoCreateInstance(CLSID_TaskScheduler))
        && SUCCEEDED(service->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t()))
        )
    {
        //std::ofstream writing_file("test.json", std::ios::out | std::ios::binary);
        //auto strbuf = std::cout.rdbuf(writing_file.rdbuf());
        if (rootPath.at(rootPath.size() - 1) == '\\')
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
        else if (std::smatch m; std::regex_match(rootPath, m, std::regex("^(.*)\\\\(.*)$")))
        {
            std::string path = m[1].str();
            std::string name = m[2].str();
            if (CComPtr<ITaskFolder> root; SUCCEEDED(service->GetFolder(_bstr_t(path.c_str()), &root)) && root)
            {
                if (CComPtr<IRegisteredTask> task; SUCCEEDED(root->GetTask(_bstr_t(name.c_str()), &task)) && task)
                {
                    std::cout << "[";
                    task_info(task, path);
                    std::cout << "]\n";
                }
            }
        }
        //std::cout.rdbuf(strbuf);
    }
    CoUninitialize();
    SetConsoleOutputCP(saveCP);
}