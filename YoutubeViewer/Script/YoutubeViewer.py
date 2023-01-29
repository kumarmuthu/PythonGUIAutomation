__version__ = "2023.01.28.01"
__author__ = "Muthukumar Subramanian"

"""
Test for Windows, Python GUI Automation script to increase the YouTube video's views
"""

import time
import argparse
import random
from collections import OrderedDict, defaultdict
import pyautogui as ms
from RuntimeBasedScriptExecution import RuntimeBasedScriptExecution
from Utility import *
import re


def argparse_function(log_obj):
    """
     ..codeauthor:: Muthukumar Subramanian
     Utility - argparse function
    :param log_obj: logger object
    :return: dict
    """
    return_dict = {}
    # Argparse
    arg_obj = argparse.ArgumentParser("Python GUI Automation to increase YouTube's views")
    arg_obj.add_argument('--main-test-runtime', type=str,
                         dest='main_test_runtime', required=False, default='1h',
                         help='Default as 1h, Run only 1 hour if argument as "1h", '
                              'Users can use Seconds, Minutes, Hours, and Days to run the test,'
                              'Example= "50s", "15m", "1h", "2d", "1y"')
    arg_obj.add_argument('--chrome-test-runtime', type=str,
                         dest='chrome_test_runtime', required=False, default='1m',
                         help='Default as 1h, Run only 1 hour if argument as "1h", '
                              'Users can use Seconds, Minutes, Hours, and Days to run the test,'
                              'Example= "50s", "15m", "1h", "2d", "1y"')
    arg_obj.add_argument('--run-config-video', type=str, choices=['true', 'false'],
                         dest='run_config_video', required=False, default='true',
                         help='Default true, Users can add more than one video links in "config.txt" file, '
                              'if false, Run hardcoded video link')
    arg_obj.add_argument('--loop-test', type=str, choices=['true', 'false'],
                         dest='loop_test', required=False, default='false',
                         help='Default false, either one or more than one video link will run one time, if true, '
                              'either one or more than one video link will repeatedly run (until main test timeout)')
    arg_obj.add_argument('--loop-count', type=str, dest='loop_count', required=False, help='based on the count')
    get_parse = arg_obj.parse_args()
    run_config_video = get_parse.run_config_video
    main_test_runtime = get_parse.main_test_runtime
    chrome_test_runtime = get_parse.chrome_test_runtime
    user_loop_test = get_parse.loop_test
    user_loop_count = get_parse.loop_count
    log_obj.info(f"run_config_video(Args Value): {run_config_video}")
    log_obj.info(f"main_test_runtime(Args Value): {main_test_runtime}")
    log_obj.info(f"chrome_test_runtime(Args Value): {chrome_test_runtime}")
    log_obj.info(f"user_loop_test(Args Value): {user_loop_test}")
    log_obj.info(f"user_loop_count(Args Value): {user_loop_count}")

    list_for_video_link = []
    default_video_link = 'https://youtu.be/jFNZsNPTSuw'
    if run_config_video == 'true':
        try:
            with open(file='config.txt', mode='r') as read_file:
                new_list = read_file.readlines()
        except FileNotFoundError as err:
            with open(file='config.txt', mode='w+') as read_file:
                read_file.write(default_video_link)
            log_obj.info("The 'config.txt' is empty!,so creating new file and adding "
                         "default video link into the config file.")
            list_for_video_link.append(default_video_link)
        else:
            log_obj.info("Reading video link from 'config.txt' file")
            list_for_video_link = new_list
    else:
        log_obj.info("Reading video link from hardcoded variable.")
        list_for_video_link.append(default_video_link)

    # Get total main test runtime in seconds format
    final_main_test_run_time = runtime_converter(log_obj, arg_obj, main_test_runtime)
    # Get chrome test runtime in seconds format
    final_chrome_test_runtime = runtime_converter(log_obj, arg_obj, chrome_test_runtime)
    # Set loop test boolean value
    if user_loop_test == 'true' and isinstance(user_loop_count, str):
        raise CustomException("Both loop-test and loop-count arguments are present!, "
                              "possible any one of the argument.")
    if user_loop_test == 'false':
        user_loop_test = False
        if isinstance(user_loop_count, str):
            try:
                user_loop_count = int(user_loop_count)
            except ValueError as err:
                raise CustomException(f"Integer is required, but used: {user_loop_count},"
                                      f"Type: {type(user_loop_count)}")
        else:
            user_loop_count = 1
    else:
        user_loop_test = True

    return_dict.update({'log_obj': log_obj, 'video_link': list_for_video_link,
                        "main_test_runtime": final_main_test_run_time,
                        'child_test_runtime': final_chrome_test_runtime,
                        'loop_test': user_loop_test,
                        'loop_count': user_loop_count})
    return return_dict


def window_session_handle(func):
    def wrapped(class_obj, *args, **kwargs):
        """
        ..codeauthor:: Muthukumar Subramanian
        Usage:
            Required argument(s):
                :param class_obj: We can use any variable here instead of 'class_obj',
                                executed class object will be present here.
                                Example: <__main__.YoutubeViewer object at 0x0000023BA70B7EF0>
                :param args: default list
                :param kwargs: default dict
                :return: YoutubeViewer.<called method|function>,
                         Example: <function YoutubeViewer.test at 0x0000023BA70B60D0>
            Optional argument(s):
                None
        """
        class_obj.log_obj.info("{:*^60}".format("From decorator function"))
        class_obj.log_obj.debug(f"Function name is: {func}")
        try:
            fw = class_obj.ms.getActiveWindow()
            class_obj.log_obj.info(f"From decorator:: Current window title name: {fw}")
            act_win = fw.isMaximized
            class_obj.log_obj.info(f"From decorator:: Check maximize: {act_win}")
            if act_win is True:
                class_obj.log_obj.info(f"From decorator:: Already maximized...")
                return True
            else:
                return func(class_obj, *args, **kwargs)
        except CustomException as err_py:
            class_obj.log_obj.error(f"Observed exception: from decorator:: {err_py}")
    return wrapped


class OptionalDecoratorManage(object):
    def __init__(self, decorator):
        """
        ..codeauthor:: Muthukumar Subramanian
        :param decorator: Actually def 'window_session_handle' will present here,
                        Example: @OptionalDecoratorManage(window_session_handle)
        """
        self.deco = decorator

    def __call__(self, func):
        """
        ..codeauthor:: Muthukumar Subramanian
        :param func: YoutubeViewer.test(.*|2|3|4)
        :return:
        """
        # self.deco(func) --> window_session_handle(YoutubeViewer.test(|2|3|4)),
        # decorator function will execute, then normal def will call
        self.deco = self.deco(func)
        self.func = func  # YoutubeViewer.test, here normal def will call

        def wrapped(*args, **kwargs):
            if kwargs.get("skip_decorator") is True:
                return self.func(*args, **kwargs)  # test will present
            else:
                return self.deco(*args, **kwargs)  # window_session_handle will present

        return wrapped


class YoutubeViewer:
    def __init__(self, *args, **kwargs):
        self.log_obj = kwargs.get("log_obj", None)
        self.chrome_job_start_time = None
        self.chrome_job_end_time = None
        self.chrome_job_runtime = 60
        self.current_active_win = None
        self.selected_title = None
        self.executed_video_count = 0
        self.total_video_count = 0
        self.post_action_executed = False
        self.job_name = kwargs.get("job_name", "paint")
        self.video_link = kwargs.get("video_link", None)
        self.ms = ms
        self.d_t = defaultdict(dict)
        self.current_position = self.ms.position()
        self.current_resolution = self.ms.size()
        # print("current_resolution: ", self.current_resolution)
        self.window_title_obj = self.pre_check_func()
        self.check_on_screen = self.ms.onScreen(146, 1048)

    def pre_check_func(self):
        """
        ..codeauthor:: Muthukumar Subramanian
        Find out window object
        :return: window title object
        """
        dict_is_there = False
        try:
            if self.window_title_obj:
                dict_is_there = True
        except AttributeError:
            pass

        _window_title_obj = defaultdict(dict)
        get_all_windows_obj = self.ms.getAllWindows()
        self.log_obj.debug("Get all the window object: {}".format(get_all_windows_obj))
        f1 = re.findall(r'\w+\d+\w+\(hWnd=\d+\)', str(get_all_windows_obj))

        list_for_all_window_title_id = []
        for i in f1:
            m1 = re.match(r'.*hWnd\=(\d+).*', str(i))
            if m1 is not None:
                list_for_all_window_title_id.append(m1.group(1))
        self.log_obj.debug(f"list_for_all_window_title_id: {list_for_all_window_title_id}")
        all_the_opened_win_title_name = self.ms.getAllTitles()
        self.log_obj.debug(f"all_the_opened_win_title_name: {all_the_opened_win_title_name}")
        temp_dict_1 = {list_for_all_window_title_id[i]: all_the_opened_win_title_name[i] for i in
                       range(len(list_for_all_window_title_id))}
        if temp_dict_1:
            if dict_is_there is False:
                for key, value in temp_dict_1.items():
                    if value == '':
                        continue
                    _window_title_obj[key] = value
                self.window_title_obj = _window_title_obj
            else:
                for key, value in temp_dict_1.items():
                    if value == '':
                        continue
                    if key not in self.window_title_obj:
                        self.window_title_obj.update({key: value})
        return self.window_title_obj

    @OptionalDecoratorManage(window_session_handle)
    def window_maximize_func(self, selected_title_name):
        """
        ..codeauthor:: Muthukumar Subramanian
        :param selected_title_name: maximizable title name
        :return: Boolean(after maximize)
        """
        try:
            time.sleep(10)
            selected_title_name.activate()
            fw = self.ms.getActiveWindow()
            self.log_obj.info(f"Current window title name: {fw}")
            act_win = fw.isMaximized
            self.log_obj.info(f"Check maximize Status!: {act_win}")
            if act_win is False:
                fw.activate()
                act = fw.isActive
                if act:
                    fw.maximize()
                    is_max = fw.isMaximized
                    if is_max:
                        self.log_obj.info("Maximized successfully-01")
                    else:
                        try:
                            self.ms.keyDown('alt')
                            self.ms.press(' ')
                            self.ms.press('x')
                            self.ms.keyUp('alt')
                            self.log_obj.info("Maximized successfully-02")
                        except CustomException as err:
                            self.log_obj.error(f"2 - Exit the script, maximized failed!, {err}")
                            raise CustomException("2 - Exit the script, maximized failed!", err)
                else:
                    try:
                        self.ms.keyDown('alt')
                        self.ms.press(' ')
                        self.ms.press('x')
                        self.ms.keyUp('alt')
                        self.log_obj.info("Maximized successfully-03")
                    except CustomException as err:
                        self.log_obj.error(f"3 - Exit the script, maximized failed!, {err}")
                        raise CustomException("3 - Exit the script, maximized failed!", err)
        except CustomException as err:
            return False
        return True

    def d_val_ret(self, d_val, ind=None):
        """
        ..codeauthor:: Muthukumar Subramanian
        For dictionary value return, based on the list's index
        :param d_val:  dict key value
        :param ind: dict value(list of index)
        :return: dict value
        """
        for key, value in self.d_t.items():
            if key == d_val:
                if ind is None:
                    return value
                return value[ind]

    def job_execute(self, exec_job, **kwargs):
        """
        ..codeauthor:: Muthukumar Subramanian
        :param exec_job: job name
        :return: Boolean
        """
        def chrome_job(**kwargs):
            self.log_obj.info("Chrome job is running...")
            random.shuffle(self.video_link)
            self.total_video_count = len(self.video_link)
            chrome_func_test_start_time = test_start_time()

            def run_chrome_job(**kwargs):
                """
                Run Chrome job
                :return: None
                """
                # Shuffle it once again for loop/count test
                random.shuffle(self.video_link)
                for ind, each_video_link in enumerate(self.video_link):
                    # Move to top search box
                    self.ms.moveTo(300, 50, duration=0)
                    self.ms.doubleClick(300, 50, interval=0)
                    # Select all text (ctrl + a)
                    self.ms.hotkey('ctrl', 'a')
                    # Clear selected text
                    self.ms.typewrite(['backspace'], interval=0)
                    # Paste video link
                    self.ms.typewrite(each_video_link, interval=0)
                    self.ms.press('enter')
                    time.sleep(15)
                    # Move cursor to end of the video
                    self.ms.moveTo(500, 540, duration=0)
                    self.ms.click(500, 540, interval=0)
                    # Move cursor a little upwards
                    self.ms.moveTo(500, 500, duration=0)
                    time.sleep(15)
                    self.executed_video_count = self.executed_video_count + 1
                # Test loop end

            try:
                calculate_executed_job_time = 0
                # Call run_chrome_job function based on the loop requirement
                if kwargs.get("loop_test"):
                    # Loop test with job timeout value
                    calculate_executed_job_time, self.total_video_count = RuntimeBasedScriptExecution(
                        self.log_obj).run_loop_test(chrome_func_test_start_time, run_chrome_job,
                                                    self.total_video_count, **kwargs)
                else:
                    # Loop count test
                    if kwargs.get("loop_count") > 1:
                        loop_cnt = kwargs.get("loop_count")
                        for i in range(0, loop_cnt):
                            run_chrome_job()
                            self.total_video_count = self.total_video_count + 1
                    else:
                        run_chrome_job()
                    end_time = test_end_time(chrome_func_test_start_time)
                    calculate_executed_job_time = (calculate_executed_job_time + int(
                        end_time.total_seconds()) - calculate_executed_job_time)
                    self.log_obj.info(f"Summary for Count test  - Job spent time(s): {calculate_executed_job_time}")
                # chrome_func_test_total_time = test_end_time(chrome_func_test_start_time)
            except CustomException as err:
                self.log_obj.error(f"Observed exception in run_chrome_job function, Exception: {err}")
                raise "Observed exception in run_chrome_job function"
            return True

        s1 = re.search(r'.*chrome.*', str(exec_job), flags=re.I)
        if s1 is not None:
            ret_bool = chrome_job(**kwargs)
            return ret_bool
        return False

    def run_test(self, *args, **kwargs):
        """
        ..codeauthor:: Muthukumar Subramanian
        GUI Automation Framework flow start here
        :param args: not used
        :param kwargs: test case related arguments
        :return: Boolean
        """
        try:
            self.ms.moveTo(146, 740, duration=0)
        except CustomException as Err:
            self.ms.FAILSAFE = False
            self.log_obj.error(f"Observed error while accessing the current window, "
                               f"Exception: {self.ms.FailSafeException}")
            self.ms.moveTo(146, 740, duration=0)
        else:
            self.log_obj.info("Successfully accessed the current window!")
        if isinstance(self.job_name, list):
            pass
        elif isinstance(self.job_name, dict) or isinstance(self.job_name, int):
            self.log_obj.error("Not yet added dict/int")
            raise CustomException("Not yet added dict/int")
        else:
            temp_list = [self.job_name]
            self.job_name = temp_list
        if self.job_name:
            for ind, each_job in enumerate(self.job_name):
                self.ms.click(x=80, y=740, interval=0)
                time.sleep(2)
                self.ms.typewrite("%s" % each_job, interval=0)
                time.sleep(2)
                self.ms.press('enter')
                time.sleep(5)
                self.window_title_obj = self.pre_check_func()
                temp_current_active_win = None
                try:
                    temp_current_active_win = self.ms.getActiveWindowTitle()
                    self.current_active_win = temp_current_active_win
                except CustomException as err:
                    self.log_obj.error(f"Observed error while getting the current active window title: {err}")
                if temp_current_active_win is None:
                    self.log_obj.warning(f"skipping the job execution for: {each_job}")
                    continue
                all_active_win_ids = []
                m2 = re.match(r'.*%s.*' % each_job, str(self.current_active_win), flags=re.I)
                if m2 is not None:
                    all_active_win_ids.extend(list(self.window_title_obj.keys()))
                else:
                    self.log_obj.warning(f"Given job name is not opened, "
                                         f"So skipping the job execution for: {each_job}, "
                                         f"Window status: {self.current_active_win}")
                    continue

                # check that app is opened or not
                self.selected_title = self.ms.Window(hWnd=int(all_active_win_ids[-1]))
                self.log_obj.info(f"selected_title: {self.selected_title}")
                ret_bool = self.window_maximize_func(self.selected_title)
                if ret_bool is False:
                    self.log_obj.warning(f"Unable to maximize the window, "
                                         f"So skipping the job execution for: {each_job}")
                    continue
                time.sleep(5)
                ret_bool_job_exec = self.job_execute(self.selected_title, **kwargs)
                if ret_bool_job_exec is False:
                    self.log_obj.error(f"Observed error while executing the job: {each_job}")
                self.post_action(each_job)
        return True

    def post_action(self, job_name=None):
        """
        ..codeauthor:: Muthukumar Subramanian
        Job post action
        :param job_name: job name
        :return: None
        """
        if job_name:
            self.log_obj.info("Closing the window for {} ...".format(job_name))
        is_active = self.selected_title.isActive
        if is_active:
            self.selected_title.close()
            self.log_obj.info("Closed the Window for job(s)")
        self.post_action_executed = True

    def run_post_test(self):
        """
        ..codeauthor:: Muthukumar Subramanian
        Main test display the video summary report
        :return: None
        """
        if self.post_action_executed is False:
            is_active = self.selected_title.isActive
            if is_active:
                self.selected_title.close()
                self.log_obj.info("Closed the Window for job(s)")
        self.log_obj.info(f"Video Summary Report:: Total video count: {self.total_video_count}, "
                          f"Executed video count: {self.executed_video_count}")


if __name__ == '__main__':
    dict_for_job = {'job_name': ['Google Chrome']}
    # Call Logger function
    get_log_obj = logger_function(__file__)
    # Call Argparse function
    ret_argparse_dict = argparse_function(get_log_obj)

    # Send all argparse variables
    dict_for_job.update(ret_argparse_dict)
    cls_obj = YoutubeViewer(**dict_for_job)
    dict_for_job.update({'run_function': 'run_test',
                         'run_post_test': True})
    utility_obj = RuntimeBasedScriptExecution(get_log_obj)
    utility_obj.main_method(dict_for_job.get('main_test_runtime'), cls_obj, **dict_for_job)
