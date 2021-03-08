'''
    History:
    Initial draft - 03.08.2021
    GUI automation supported windows system
        I - width= 1920, height= 1080
       II - width=1366, height=768
'''


__version__ = "2021.03.08.01"
__author__ = "Muthukumar Subramanian"


import re
import time
from collections import OrderedDict, defaultdict
import pyautogui as ms


def windowSessionHandle(func):
    def wrapped(class_obj, *args, **kwargs):
        '''
        ..codeauthor:: Muthukumar Subramanian
        Usage:
            Required argument(s):
                :param class_obj: We can use any variable here instead of 'class_obj',
                                executed class object will be present here.
                                Example: <__main__.MuthuGUIAutomation object at 0x0000023BA70B7EF0>
                :param args: default list
                :param kwargs: default dict
                :return: MuthuGUIAutomation.<called method|function>,
                         Example: <function MuthuGUIAutomation.test at 0x0000023BA70B60D0>
            Optional argument(s):
                None
        '''
        print("From decorator function".format())
        print("Function name is: {}".format(func))
        try:
            fw = class_obj.ms.getActiveWindow()
            print("from decorator:: Current window title name: ", fw)
            act_win = fw.isMaximized
            print("from decorator:: Check maximize: ", act_win)
            if act_win is True:
                print("from decorator:: Already maximized...".format())
                return True
            else:
                return func(class_obj, *args, **kwargs)
        except Exception as err_py:
            print("Observed exception: from decorator:: {}".format(err_py))
    return wrapped


class OptionalDecoratorManage(object):
    def __init__(self, decorator):
        '''
        ..codeauthor:: Muthukumar Subramanian
        :param decorator: Actually def 'windowSessionHandle' will present here,
                        Example: @OptionalDecoratorManage(windowSessionHandle)
        '''
        self.deco = decorator

    def __call__(self, func):
        '''
        ..codeauthor:: Muthukumar Subramanian
        :param func: MuthuGUIAutomation.test(.*|2|3|4)
        :return:
        '''
        # self.deco(func) --> windowSessionHandle(MuthuGUIAutomation.test(|2|3|4)),
        # decorator function will execute, then normal def will call
        self.deco = self.deco(func)
        self.func = func  # MuthuGUIAutomation.test, here normal def will call

        def wrapped(*args, **kwargs):
            if kwargs.get("skip_decorator") is True:
                return self.func(*args, **kwargs)  # test will present
            else:
                return self.deco(*args, **kwargs)  # windowSessionHandle will present

        return wrapped


class MuthuGUIAutomation():
    def __init__(self, *args, **kwargs):
        self.current_active_win = None
        self.ms = ms
        self.d_t = defaultdict(dict)
        self.current_position = self.ms.position()
        self.current_resolution = self.ms.size()
        # print("current_resolution: ", self.current_resolution)
        self.window_title_obj = self.pre_check_func()
        self.job_name = kwargs.get("job_name", "paint")
        # self.check_on_screen = self.ms.onScreen(146, 1048)

    def pre_check_func(self):
        dict_is_there = False
        try:
            if self.window_title_obj:
                dict_is_there = True
        except AttributeError:
            pass

        _window_title_obj = defaultdict(dict)
        get_all_windows_obj = self.ms.getAllWindows()
        print("Get all the window object: {}".format(get_all_windows_obj))
        f1 = re.findall(r'\w+\d+\w+\(hWnd=\d+\)', str(get_all_windows_obj))

        list_for_all_window_title_id = []
        for i in f1:
            m1 = re.match(r'.*hWnd\=(\d+).*', str(i))
            if m1 is not None:
                list_for_all_window_title_id.append(m1.group(1))
        print("list_for_all_window_title_id: ", list_for_all_window_title_id)
        all_the_opened_win_title_name = self.ms.getAllTitles()
        print("all_the_opened_win_title_name: ", all_the_opened_win_title_name)
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

    @OptionalDecoratorManage(windowSessionHandle)
    def window_maximize_func(self, selected_title_name):
        try:
            time.sleep(10)
            selected_title_name.activate()
            fw = self.ms.getActiveWindow()
            print("Current window title name: ", fw)
            act_win = fw.isMaximized
            print("Check maximize!", act_win)
            if act_win is False:
                fw.activate()
                act = fw.isActive
                if act:
                    fw.maximize()
                    is_max = fw.isMaximized
                    if is_max:
                        print("Maximized successfully-01")
                    else:
                        try:
                            self.ms.keyDown('alt')
                            self.ms.press(' ')
                            self.ms.press('x')
                            self.ms.keyUp('alt')
                            print("Maximized successfully-02")
                        except Exception as err:
                            raise Exception("2 - Exit the script, maximized failed!", err)
                else:
                    try:
                        self.ms.keyDown('alt')
                        self.ms.press(' ')
                        self.ms.press('x')
                        self.ms.keyUp('alt')
                        print("Maximized successfully-03")
                    except Exception as err:
                        raise Exception("3 - Exit the script, maximized failed!", err)
        except Exception as err:
            return False
        return True

    def d_val_ret(self, d_val, ind=None):
        for key, value in self.d_t.items():
            if key == d_val:
                if ind is None:
                    return value
                return value[ind]

    def job_execute(self, exec_job):
        # paint job
        def paint_job():
            print("Paint job is running...")
            if re.match(r'.*1920.*1080.*', str(self.current_resolution)):
                self.d_t = {'_default_duration': 0.2, 'resolution': [1920, 1080],
                            'distance': 300, 'change': 20, 'rose_color': [1200, 120],
                            'move_to_1': [300, 400], 'move_right': [120, 0],
                            'move_down': [0, 120], 'move_left': [-120, 0],
                            'move_up': [0, -120], 'green_color': [1300, 120],
                            'move_to_2': [400, 500], 'blue_color': [1350, 120],
                            'move_to_3': [500, 600], 'file_select': [40, 51],
                            'file_save': [40, 220], 'move_to_4': [1200, 655],
                            'save_as_double_click': [1200, 655], 'save_click': [1200, 758],
                            'save_cancel': [1310, 758], 'over_write': [1010, 516]
                            }
            elif re.match(r'.*1366.*768.*', str(self.current_resolution)):
                self.d_t = {'resize': [180, 80], 'select_pixel': [220, 152],
                            'horizontal_text': [254, 183], 'click_ok': [180, 450],
                            '_default_duration': 0.2, 'resolution': [1366, 768],
                            'distance': 300, 'change': 20, 'rose_color': [830, 80],
                            'move_to_1': [100, 200], 'move_right': [120, 0],
                            'move_down': [0, 120], 'move_left': [-120, 0],
                            'move_up': [0, -120], 'green_color': [895, 80],
                            'move_to_2': [200, 300], 'blue_color': [935, 80],
                            'move_to_3': [300, 400], 'file_select': [38, 35],
                            'file_save': [38, 165], 'move_to_4': [520, 410],
                            'save_as_double_click': [520, 410], 'save_click': [520, 470],
                            'save_cancel': [570, 470], 'over_write': [730, 372],
                            'click_exit': [1330, 10], 'click_again': [625, 390]
                            }
            try:
                for i in range(1):
                    self.ms.moveTo(self.d_val_ret('resize', 0), self.d_val_ret('resize', 1), duration=0.2)  # resize
                    self.ms.click(self.d_val_ret('resize', 0), self.d_val_ret('resize', 1), interval=0.2)
                    self.ms.moveTo(self.d_val_ret('select_pixel', 0), self.d_val_ret('select_pixel', 1),
                                   duration=0.2)  # select pixel
                    self.ms.click(self.d_val_ret('select_pixel', 0), self.d_val_ret('select_pixel', 1), interval=0.2)
                    self.ms.moveTo(self.d_val_ret('horizontal_text', 0), self.d_val_ret('horizontal_text', 1),
                                   duration=0.2)  # horizontal
                    self.ms.doubleClick(
                        self.d_val_ret(
                            'horizontal_text', 0), self.d_val_ret(
                            'horizontal_text', 1), interval=0)
                    self.ms.typewrite(['backspace'], interval=0)
                    self.ms.typewrite(str(self.d_val_ret('resolution', 0)), interval=0)
                    self.ms.moveTo(self.d_val_ret('click_ok', 0), self.d_val_ret('click_ok', 1), duration=0.2)  # ok
                    self.ms.click(self.d_val_ret('click_ok', 0), self.d_val_ret('click_ok', 1), interval=0.2)

                    # 1st square
                    self.ms.moveTo(
                        self.d_val_ret(
                            'rose_color', 0), self.d_val_ret(
                            'rose_color', 1), duration=0.2)  # rose
                    self.ms.click(self.d_val_ret('rose_color', 0), self.d_val_ret('rose_color', 1), interval=0.2)
                    self.ms.moveTo(self.d_val_ret('move_to_1', 0), self.d_val_ret('move_to_1', 1), duration=0.2)
                    # Move right.
                    self.ms.drag(self.d_val_ret('move_right', 0), self.d_val_ret('move_right', 1), duration=0.2)
                    # Move down.
                    self.ms.drag(self.d_val_ret('move_down', 0), self.d_val_ret('move_down', 1), duration=0.2)
                    # Move left.
                    self.ms.drag(self.d_val_ret('move_left', 0), self.d_val_ret('move_left', 1), duration=0.2)
                    self.ms.drag(self.d_val_ret('move_up', 0), self.d_val_ret('move_up', 1), duration=0.2)  # Move up.

                    # 2nd square
                    self.ms.moveTo(
                        self.d_val_ret(
                            'green_color', 0), self.d_val_ret(
                            'green_color', 1), duration=0.2)  # green
                    self.ms.click(self.d_val_ret('green_color', 0), self.d_val_ret('green_color', 1), interval=0.2)
                    self.ms.moveTo(self.d_val_ret('move_to_2', 0), self.d_val_ret('move_to_2', 1), duration=0)
                    # Move right.
                    self.ms.drag(self.d_val_ret('move_right', 0), self.d_val_ret('move_right', 1), duration=0.2)
                    # Move down.
                    self.ms.drag(self.d_val_ret('move_down', 0), self.d_val_ret('move_down', 1), duration=0.2)
                    # Move left.
                    self.ms.drag(self.d_val_ret('move_left', 0), self.d_val_ret('move_left', 1), duration=0.2)
                    self.ms.drag(self.d_val_ret('move_up', 0), self.d_val_ret('move_up', 1), duration=0.2)  # Move up.

                    # 3rd square
                    self.ms.moveTo(
                        self.d_val_ret(
                            'blue_color', 0), self.d_val_ret(
                            'blue_color', 1), duration=0.2)  # blue
                    self.ms.click(self.d_val_ret('blue_color', 0), self.d_val_ret('blue_color', 1), interval=0.2)
                    self.ms.moveTo(self.d_val_ret('move_to_3', 0), self.d_val_ret('move_to_3', 1), duration=0.2)
                    distance = self.d_val_ret('distance')
                    change = self.d_val_ret('change')
                    while distance > 0:
                        self.ms.drag(distance, 0, duration=0.2)  # Move right.
                        distance = distance - change
                        self.ms.drag(0, distance, duration=0.2)  # Move down.
                        self.ms.drag(-distance, 0, duration=0.2)  # Move left.
                        distance = distance - change
                        self.ms.drag(0, -distance, duration=0.2)  # Move up.

                    self.ms.moveTo(
                        self.d_val_ret(
                            'file_select', 0), self.d_val_ret(
                            'file_select', 1), duration=0.2)  # File
                    self.ms.click(self.d_val_ret('file_select', 0), self.d_val_ret('file_select', 1), interval=0)
                    self.ms.moveTo(self.d_val_ret('file_save', 0), self.d_val_ret('file_save', 1), duration=0.2)  # Save
                    self.ms.click(self.d_val_ret('file_save', 0), self.d_val_ret('file_save', 1), interval=0)
                    self.ms.moveTo(self.d_val_ret('move_to_4', 0), self.d_val_ret('move_to_4', 1), duration=0.2)
                    self.window_title_obj = self.pre_check_func()
                    self.ms.doubleClick(
                        x=self.d_val_ret(
                            'save_as_double_click', 0), y=self.d_val_ret(
                            'save_as_double_click', 0), interval=0)
                    self.ms.typewrite(['backspace'], interval=0)
                    self.ms.typewrite('Muthu', interval=0)
                    self.ms.moveTo(
                        self.d_val_ret(
                            'save_click', 0), self.d_val_ret(
                            'save_click', 1), duration=0.2)  # Save
                    self.ms.click(self.d_val_ret('save_click', 0), self.d_val_ret('save_click', 1), interval=0)
                    self.ms.moveTo(
                        self.d_val_ret(
                            'save_cancel', 0), self.d_val_ret(
                            'save_cancel', 1), duration=0.2)  # Cancel
                    self.ms.click(self.d_val_ret('save_cancel', 0), self.d_val_ret('save_cancel', 1), interval=0)
                    self.ms.moveTo(
                        self.d_val_ret(
                            'over_write', 0), self.d_val_ret(
                            'over_write', 1), duration=0.2)  # Over write
                    self.ms.doubleClick(
                        x=self.d_val_ret(
                            'over_write', 0), y=self.d_val_ret(
                            'over_write', 1), interval=0)
                    self.ms.moveTo(
                        self.d_val_ret(
                            'click_exit', 0), self.d_val_ret(
                            'click_exit', 1), duration=0.2)  # exit
                    self.ms.click(self.d_val_ret('click_exit', 0), self.d_val_ret('click_exit', 1), interval=0)
                    self.ms.moveTo(
                        self.d_val_ret(
                            'click_again', 0), self.d_val_ret(
                            'click_again', 1), duration=0.2)  # save again
                    self.ms.click(self.d_val_ret('click_again', 0), self.d_val_ret('click_again', 1), interval=0)
                    time.sleep(1)
            except Exception as err:
                print('Observed error while closing the current window:- ', err)
                return False
            return True

        def chrome_job():
            print("Chrome job is running...")
            self.ms.moveTo(1200, 600, duration=0.2)
            self.ms.click(1200, 600, interval=0.2, button='right')  # right click
            return True

        s1 = re.search(r'.*paint.*', str(exec_job), flags=re.I)
        s2 = re.search(r'.*chrome.*', str(exec_job), flags=re.I)
        if s1 is not None:
            ret_paint = paint_job()
            return ret_paint
        if s2 is not None:
            ret_chrome = chrome_job()
            return ret_chrome
        return False

    def run_test(self, *args, **kwargs):
        try:
            self.ms.moveTo(146, 1048, duration=0.25)
        except Exception as Err:
            self.ms.FAILSAFE = False
            print("Observed error while accessing the current window", self.ms.FailSafeException)
            self.ms.moveTo(146, 1048, duration=0.25)
        else:
            print("Successfully accessed the current window!")

        if isinstance(self.job_name, list):
            pass
        elif isinstance(self.job_name, dict) or isinstance(self.job_name, int):
            raise Exception("Not yet added dict/int")
        else:
            temp_list = [self.job_name]
            self.job_name = temp_list
        if self.job_name:
            for each_job in self.job_name:
                self.ms.click(x=146, y=1048, interval=0)
                self.ms.typewrite("%s" % each_job, interval=0)
                time.sleep(2)
                self.ms.press('enter')
                time.sleep(5)
                self.window_title_obj = self.pre_check_func()
                temp_current_active_win = None
                try:
                    temp_current_active_win = self.ms.getActiveWindowTitle()
                    self.current_active_win = temp_current_active_win
                except Exception as err:
                    print("Observed error while getting the current active window title: ", err)
                if temp_current_active_win is None:
                    print("skipping the job execution for: {}".format(each_job))
                    continue
                all_active_win_ids = []
                m2 = re.match(r'.*%s.*' % each_job, str(self.current_active_win), flags=re.I)
                if m2 is not None:
                    all_active_win_ids = list(self.window_title_obj.keys())
                else:
                    print(each_job, self.current_active_win)
                    print("Given job name is not opened, so skipping the job execution for: {}".format(each_job))
                    continue

                # check that app is opened or not
                selected_title = self.ms.Window(hWnd=int(all_active_win_ids[-1]))
                print("selected_title: ", selected_title)
                ret_bool = self.window_maximize_func(selected_title)
                if ret_bool is False:
                    print("Unable to maximize the window, so skipping the job execution for: {}".format(each_job))
                    continue
                time.sleep(5)
                ret_bool_job_exec = self.job_execute(selected_title)
                if ret_bool_job_exec is False:
                    print("Observed error while executing the job: ", each_job)
                print("Closing the window for {} ...".format(each_job))
                is_active = selected_title.isActive
                if is_active:
                    selected_title.close()
        return True


if __name__ == '__main__':
    dict_for_job = {}
    dict_for_job.update({'job_name': ['Paint', 'Google Chrome']})
    cls_obj = MuthuGUIAutomation(**dict_for_job)
    cls_obj.run_test()
