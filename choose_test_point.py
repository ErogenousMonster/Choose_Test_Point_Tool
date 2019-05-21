# coding=utf-8
import xlwings as xw
import os
import re
import difflib
import copy


def classification_signal_line(net_dict, net_type=None):
    """将信号线按组分类，每类取一个（组）信号线"""
    new_net_dict = {}
    for key, value in net_dict.items():
        new_value = []
        value_list = copy.deepcopy(value)
        # print('')
        # print(value_list)
        while value_list:
            if net_type == 'diff':
                value_0 = copy.deepcopy(value_list[0])
                new_value.append(value_list[0])
                new_value.append(value_list[1])
                for i in range(0, len(value), 2):
                    if difflib.SequenceMatcher(None, value_0, value[i]).quick_ratio() >= 0.6:
                        try:
                            value_list.remove(value[i])
                            value_list.remove(value[i + 1])
                        except:
                            pass

            if net_type == 'single':
                value_0 = copy.deepcopy(value_list[0])
                new_value.append(value_list[0])
                for i in range(0, len(value)):
                    if difflib.SequenceMatcher(None, value_0, value[i]).quick_ratio() >= 0.6:
                        try:
                            value_list.remove(value[i])
                        except:
                            pass
        if net_type == 'single':
            # 单根取四根线
            new_net_dict[key] = new_value[:4] if len(new_value) > 4 else new_value
        else:
            # 差分取4组线
            new_net_dict[key] = new_value[-8:] if len(new_value) > 8 else new_value

    return new_net_dict


# 设置单元格内容的字型字体大小和字体位置
def SetCellFont(sheet, cell_ind, Font_Name, Font_Size, Border_size=1, Font_Bold=False, horizon_alignment='c'):
    sheet.range(cell_ind).api.Font.Name = Font_Name
    sheet.range(cell_ind).api.Font.Size = Font_Size
    sheet.range(cell_ind).api.Font.Bold = Font_Bold
    sheet.range(cell_ind).current_region.api.Borders.LineStyle = Border_size
    if horizon_alignment == 'c':
        sheet.range(cell_ind).api.HorizontalAlignment = -4108
    elif horizon_alignment == 'r':
        sheet.range(cell_ind).api.HorizontalAlignment = -4152
    elif horizon_alignment == 'l':
        sheet.range(cell_ind).api.HorizontalAlignment = -4131


class ChooseTestPoint:
    def __init__(self):
        self.checklist_path = None
        self.brd_path = None
        self.output_path = None

        # get_net_list_from_checklist
        self.diff_net_list_from_checklist = []
        self.single_net_list = []

        # export_brd_report
        self.diff_net_list = []
        self.diff_net_list_from_report = []
        self.diff_pair_spacing_dict = {}
        self.diff_pair_one_spacing_dict = {}
        self.diff_net_one_spacing_dict = {}
        self.net_layer_dict = {}
        self.net_width_dict = {}
        self.diff_pair_spacing_length_dict = {}
        self.diff_net_one_spacing_length_dict = {}
        self.npr_net_diff_pair_dict = {}
        self.npr_diff_pair_net_dict = {}

        # get_all_specifications_from_output_file
        self.outer_single_width_list = []
        self.outer_diff_width_list = []
        self.outer_diff_spacing_list = []
        self.outer_ws_impedance_dict = {}
        self.inner_single_width_list = []
        self.inner_diff_width_list = []
        self.inner_diff_spacing_list = []
        self.inner_ws_impedance_dict = {}

        # get_suitable_net
        self.outer_single_width_net_dict = {}
        self.inner_single_width_net_dict = {}
        self.outer_diff_width_net_dict = {}
        self.inner_diff_width_net_dict = {}

        self.root_path = os.getcwd()
        for x in os.listdir(self.root_path):
            # checklist
            if x.find('.xlsm') > -1:
                self.checklist_path = os.path.join(self.root_path, x)
            elif x.find('.brd') > -1:
                self.brd_path = os.path.join(self.root_path, x)
            elif x.find('.xls') > -1 or x.find('.xlsx') > -1:
                self.output_path = os.path.join(self.root_path, x)

        # 异常处理：如果需要输入不存在
        if self.checklist_path is None:
            print('Checklist file does not exist.Please check!')
            os.system("pause")
            raise FileNotFoundError

        if self.brd_path is None:
            print('Brd file does not exist.Please check!')
            os.system("pause")
            raise FileNotFoundError

        if self.output_path is None:
            print('Output file does not exist.Please check!')
            os.system("pause")
            raise FileNotFoundError

    def get_net_list_from_checklist(self):
        """从checklist的netlist表中读出最终的差分和单根信号线列表"""

        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(self.checklist_path)
        sht = wb.sheets['Netlist']

        diff_ind, single_ind = None, None

        for cell in sht.api.UsedRange.Cells:
            if cell.Value == 'Differential':
                diff_ind = (cell.Row + 1, cell.Column)
            elif cell.Value == 'Single-Ended':
                single_ind = (cell.Row + 1, cell.Column)
                break

        # 异常处理
        if diff_ind is None:
            print('The Netlist sheet in the checklist excel has no differential signal lines.')
            os.system("pause")
            raise FileNotFoundError

        if single_ind is None:
            print('The Netlist sheet in the checklist excel has no single signal lines.')
            os.system("pause")
            raise FileNotFoundError

        # 这里的差分信号线是所有的，后面要和allegro生成的对比进行筛选
        self.diff_net_list_from_checklist = sum(sht.range(diff_ind).options(expand='table', ndim=2).value, [])
        self.single_net_list = sht.range(single_ind).options(expand='table', ndim=1).value

        wb.close()
        app.quit()

    def get_all_specifications_from_output_file(self):
        """从output_file中读出target impedance"""
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(self.output_path)
        sht = wb.sheets['疊構&阻抗 Requirement']

        target_impedance_ind = None

        for cell in sht.api.UsedRange.Cells:
            if cell.Value == 'Trace width/Spacing (mil)':
                target_impedance_ind = (cell.Row + 2, cell.Column)
                break

        # 异常处理
        if target_impedance_ind is None:
            print('The 疊構&阻抗 Requirement sheet in the output excel has no target impedance table.')
            os.system("pause")
            raise FileNotFoundError

        # 外层信号线信息
        outer_trace_width_spacing = sht.range(target_impedance_ind).options(expand='table', ndim=1).value
        outer_trace_layer = sht.range((target_impedance_ind[0], target_impedance_ind[1] + 2))\
            .options(expand='table', ndim=1).value
        outer_trace_type = sht.range((target_impedance_ind[0], target_impedance_ind[1] + 4)) \
            .options(expand='table', ndim=1).value
        # 内层信号线信息
        inner_trace_width_spacing = sht.range((target_impedance_ind[0], target_impedance_ind[1] + 6))\
            .options(expand='table', ndim=1).value
        inner_trace_layer = sht.range((target_impedance_ind[0], target_impedance_ind[1] + 8))\
            .options(expand='table', ndim=1).value
        inner_trace_type = sht.range((target_impedance_ind[0], target_impedance_ind[1] + 10)) \
            .options(expand='table', ndim=1).value

        # 对target impedance信息进行处理
        outer_length = len(outer_trace_type)
        outer_trace_width_spacing = outer_trace_width_spacing[:outer_length]
        outer_trace_layer = outer_trace_layer[:outer_length]

        inner_length = len(inner_trace_type)
        inner_trace_width_spacing = inner_trace_width_spacing[:inner_length]
        inner_trace_layer = inner_trace_layer[:inner_length]

        for outer_i in range(outer_length):
            if outer_trace_type[outer_i] == 'Single-ended':
                outer_width_item = float(outer_trace_width_spacing[outer_i])
                self.outer_single_width_list.append(str('%.2f' % outer_width_item))
                self.outer_ws_impedance_dict[str('%.2f' % outer_width_item)] = outer_trace_layer[outer_i]
            else:
                outer_width_item, outer_spacing_item, _ = outer_trace_width_spacing[outer_i].split('/')
                outer_width_item = float(outer_width_item)
                outer_spacing_item = float(outer_spacing_item)
                self.outer_diff_width_list.append(outer_width_item)
                self.outer_diff_spacing_list.append(outer_spacing_item)
                self.outer_ws_impedance_dict[str('%.2f' % outer_width_item) + ' ' + str('%.2f' % outer_spacing_item)] \
                    = outer_trace_layer[outer_i]

        for inner_i in range(inner_length):
            if inner_trace_type[inner_i] == 'Single-ended':
                inner_width_item = float(inner_trace_width_spacing[inner_i])
                self.inner_single_width_list.append(str('%.2f' % inner_width_item))
                self.inner_ws_impedance_dict[str('%.2f' % inner_width_item)] = inner_trace_layer[inner_i]
            else:
                inner_width_item, inner_spacing_item, _ = inner_trace_width_spacing[inner_i].split('/')
                inner_width_item = float(inner_width_item)
                inner_spacing_item = float(inner_spacing_item)
                self.inner_diff_width_list.append(inner_width_item)
                self.inner_diff_spacing_list.append(inner_spacing_item)
                self.inner_ws_impedance_dict[str('%.2f' % inner_width_item) + ' ' + str('%.2f' % inner_width_item)] \
                    = inner_trace_layer[inner_i]

        wb.close()
        app.quit()

    def export_brd_report(self):
        """从brd文件中生成 Etch Detailed Length 及 Diffpair Gaps 及 Properties on Nets Report并分析数据"""
        report_list = ['elw.rpt', 'dpg.rpt', 'npr.rpt']
        command_list = ['elw', 'dpg', 'npr']

        complete_command_list = ['report -v %s "%s" "%s"' % (command_list[i], self.brd_path,
                                                             os.path.join(self.root_path, report_list[i]))
                                 for i in range(len(command_list))]

        # 生成报告的路径
        elw_path = os.path.join(self.root_path, 'elw.rpt')
        dpg_path = os.path.join(self.root_path, 'dpg.rpt')
        npr_path = os.path.join(self.root_path, 'npr.rpt')

        # 如果有之前生成的报告则先删除
        if os.path.exists(elw_path):
            os.remove(elw_path)
        if os.path.exists(dpg_path):
            os.remove(dpg_path)
        if os.path.exists(npr_path):
            os.remove(npr_path)

        # 指令生成报告
        for command in complete_command_list:
            os.system(command)

        # 读取eld.rpt报告中的数据
        # 筛选出有diff pair的差分信号线
        with open(npr_path) as npr_file:
            npr_content = npr_file.read().split('\n')[5:-1]

            for npr_item in npr_content:
                npr_item_list = npr_item.split(',')
                # 如果diff pair有值
                if npr_item_list[25]:
                    diff_net = self.npr_diff_pair_net_dict.get(npr_item_list[25])
                    self.npr_net_diff_pair_dict[npr_item_list[0]] = npr_item_list[25]
                    self.npr_diff_pair_net_dict[npr_item_list[25]] = diff_net + [npr_item_list[0]] \
                        if diff_net else [npr_item_list[0]]

        # 排除没有diff pair的差分信号线
        self.diff_net_list = [i for i in self.diff_net_list_from_checklist
                              if i in self.npr_net_diff_pair_dict.keys()]
        # print('diff', self.diff_net_list)
        # print('')

        # 读取elw.rpt报告中的数据
        net_length_dict = {}
        with open(elw_path) as elw_file:
            elw_content = elw_file.read().split('\n')[5:-1]
            # print(elw_content)

            for elw_item in elw_content:
                elw_item_list = elw_item.split(',')
                # 只有长度占比大于60.00%并且自身长度大于1000mil才考虑进来
                total_length = float(elw_item_list[2])
                layer_length = float(elw_item_list[4])
                net_name = elw_item_list[0]
                net_length_dict[net_name] = total_length

                if (layer_length / total_length) >= 0.60 and layer_length >= 1000.00:

                    net_layer = elw_item_list[1]
                    net_width = float(elw_item_list[3])

                    self.net_layer_dict[net_name] = self.net_layer_dict.get(net_name) \
                        if self.net_layer_dict.get(net_name) else net_layer
                    self.net_width_dict[net_name] = self.net_width_dict.get(net_name) \
                        if self.net_width_dict.get(net_name) else str('%.2f' % net_width)

        # 读取dpg.rpt报告中的数据
        with open(dpg_path) as dpg_file:
            dpg_content = dpg_file.read().split('\n')[5:-1]

            for dpg_item in dpg_content:
                dpg_item_list = dpg_item.split(',')
                dpg_diff_pair = dpg_item_list[0].split(' ')[0][1:]
                dpg_spacing = dpg_item_list[2]
                dpg_length = float(dpg_item_list[-2])

                # 保存差分信号线组的间距（同一组差分线的线间距理论上只有一个，但是实际可能有多个）
                diff_pair_spacing = self.diff_pair_spacing_dict.get(dpg_diff_pair)
                if diff_pair_spacing is not None:
                    # 不为None说明里面有内容
                    if dpg_spacing == diff_pair_spacing[-1]:
                        # 如果此间距和上间距相同，则合并
                        self.diff_pair_spacing_length_dict[dpg_diff_pair][-1] = \
                            self.diff_pair_spacing_length_dict.get(dpg_diff_pair)[-1] + dpg_length
                    else:
                        self.diff_pair_spacing_length_dict[dpg_diff_pair] = \
                            self.diff_pair_spacing_length_dict.get(dpg_diff_pair) + [dpg_length]
                        # 如果间距不存在，就添加进去
                        self.diff_pair_spacing_dict[dpg_diff_pair] = diff_pair_spacing + [dpg_spacing]

                else:
                    self.diff_pair_spacing_length_dict[dpg_diff_pair] = [dpg_length]
                    self.diff_pair_spacing_dict[dpg_diff_pair] = [dpg_spacing]

        # 对两种及多种spacing的情况进行讨论,取出最长的那个
        for key, value in self.diff_pair_spacing_dict.items():
            # 如果有两种及两种以上的情况
            if len(value) > 1:
                length_list = self.diff_pair_spacing_length_dict[key]
                longest_index = length_list.index(max(length_list))
                self.diff_pair_one_spacing_dict[key] = value[longest_index]
            else:
                self.diff_pair_one_spacing_dict[key] = value[0]

        # print('diff_pair_one_spacing_dict', self.diff_pair_one_spacing_dict)
        # print('')
        # print('npr_diff_pair_net_dict', self.npr_diff_pair_net_dict)
        # print('')

        # 将key值从diff pair改为最长的差分信号线组
        for diff_key, diff_value in self.diff_pair_one_spacing_dict.items():
            # 如果diff pair存在
            if self.npr_diff_pair_net_dict.get(diff_key):
                # 如果长为2，说明只有一组差分对(self.npr_diff_pair_net_dict一定有值，最少为一对差分对)
                net_list = self.npr_diff_pair_net_dict[diff_key]
                if len(net_list) == 2:
                    self.diff_net_one_spacing_dict[net_list[0]] = diff_value
                    self.diff_net_one_spacing_dict[net_list[1]] = diff_value
                # 含有多对差分对，找出最长的那对
                else:
                    net_copy_list = copy.deepcopy(net_list)
                    net_reordered_list = []
                    # 将 net_copy_list的数据重新排序
                    while net_copy_list:
                        net_1 = net_copy_list[0]
                        try:
                            net_2 = self.diff_net_list[self.diff_net_list.index(net_1) + 1]
                            net_reordered_list.append(net_1)
                            net_reordered_list.append(net_2)
                            net_copy_list.remove(net_1)
                            net_copy_list.remove(net_2)
                        except ValueError:
                            net_reordered_list = []
                            break
                    if net_reordered_list:
                        # if diff_key == 'USB3_TX2_D2':
                        #     print(1, net_reordered_list)
                        net_len_list = []
                        for i in range(0, len(net_reordered_list), 2):
                            net_len_list.append(net_length_dict[net_reordered_list[i]])
                        max_ind = net_len_list.index(max(net_len_list)) * 2
                        # if diff_key == 'USB3_TX2_D2':
                        #     print(net_len_list)
                        #     print(max_ind)
                        self.diff_net_one_spacing_dict[net_reordered_list[max_ind]] = diff_value
                        self.diff_net_one_spacing_dict[net_reordered_list[max_ind + 1]] = diff_value

        # print(self.diff_net_one_spacing_dict)
        # print('')

    def _get_suitable_net(self):
        """找出符合规则的信号线及其信息"""
        # 先找出单根线中符合规则的net
        for single_net_item in self.single_net_list:
            for single_width in self.outer_single_width_list:
                single_layer = self.net_layer_dict.get(single_net_item)
                net_width = self.net_width_dict.get(single_net_item)
                if single_layer and net_width and single_layer in ['TOP', 'BOTTOM'] and single_width == net_width:
                    single_width = str(single_width)
                    single_key = single_width + ' ' + single_layer
                    self.outer_single_width_net_dict[single_key] = \
                        self.outer_single_width_net_dict.get(single_key) + [single_net_item] if \
                        self.outer_single_width_net_dict.get(single_key) else [single_net_item]

            for single_width in self.inner_single_width_list:
                single_layer = self.net_layer_dict.get(single_net_item)
                net_width = self.net_width_dict.get(single_net_item)
                # if single_net_item == 'M_B_MA6':
                #     print(single_layer, net_width)
                if single_layer and net_width and single_layer not in ['TOP', 'BOTTOM'] and single_width == net_width:
                    single_width = str(single_width)
                    single_key = single_width + ' ' + single_layer
                    self.inner_single_width_net_dict[single_key] = \
                        self.inner_single_width_net_dict.get(single_key) + [single_net_item] if \
                        self.inner_single_width_net_dict.get(single_key) else [single_net_item]

        # 找出差分符合规则的信号线
        outer_ws_list = list(self.outer_ws_impedance_dict.keys())
        inner_ws_list = list(self.inner_ws_impedance_dict.keys())

        for diff_net_item in self.diff_net_list:
            for diff_width_spacing in outer_ws_list:
                diff_layer = self.net_layer_dict.get(diff_net_item)
                net_width = self.net_width_dict.get(diff_net_item)
                net_spacing = self.diff_net_one_spacing_dict.get(diff_net_item)
                if diff_layer and net_width and net_spacing:
                    net_width = str('%.2f' % float(net_width))
                    net_spacing = str('%.2f' % float(net_spacing))
                    net_width_spacing = net_width + ' ' + net_spacing
                    diff_key = net_width_spacing + ' ' + diff_layer
                    if diff_layer in ['TOP', 'BOTTOM'] and net_width_spacing == diff_width_spacing:
                        self.outer_diff_width_net_dict[diff_key] = \
                            self.outer_diff_width_net_dict.get(diff_key) + [diff_net_item] if \
                                self.outer_diff_width_net_dict.get(diff_key) else [diff_net_item]

            for diff_width_spacing in inner_ws_list:
                diff_layer = self.net_layer_dict.get(diff_net_item)
                net_width = self.net_width_dict.get(diff_net_item)
                net_spacing = self.diff_net_one_spacing_dict.get(diff_net_item)
                if diff_layer and net_width and net_spacing:
                    net_width = str('%.2f' % float(net_width))
                    net_spacing = str('%.2f' % float(net_spacing))
                    net_width_spacing = net_width + ' ' + net_spacing
                    diff_key = net_width_spacing + ' ' + diff_layer
                    if diff_layer not in ['TOP', 'BOTTOM'] and net_width_spacing == diff_width_spacing:
                        self.inner_diff_width_net_dict[diff_key] = \
                            self.inner_diff_width_net_dict.get(diff_key) + [diff_net_item] if \
                                self.inner_diff_width_net_dict.get(diff_key) else [diff_net_item]
        # print(self.outer_single_width_net_dict)
        # print('')
        # print(self.inner_single_width_net_dict)
        # print('')
        # print(self.outer_diff_width_net_dict)
        # print('')
        # print(self.inner_diff_width_net_dict)
        # print('')
        return classification_signal_line(self.outer_single_width_net_dict, net_type='single'), \
               classification_signal_line(self.inner_single_width_net_dict, net_type='single'), \
               classification_signal_line(self.outer_diff_width_net_dict, net_type='diff'), \
               classification_signal_line(self.inner_diff_width_net_dict, net_type='diff')

    def create_output_file(self):
        part_outer_single_dict, part_inner_single_dict, part_outer_diff_dict, part_inner_diff_dict \
            = self._get_suitable_net()

        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(self.output_path)
        try:
            sht = wb.sheets['疊構&阻抗 Requirement']
        except:
            print("Sheet '疊構&阻抗 Requirement' does not exist.Please check!")
            os.system("pause")
            raise FileNotFoundError

        table_ind = None

        # 找到要生成表格的起始坐标
        for cell in sht.api.UsedRange.Cells:
            if cell.Value == '2.2 板廠的QC報告應包含這些測試線的阻抗驗證結果':
                table_ind = (cell.Row + 3, cell.Column)
                break

        # 生成填入表格的数据
        output_list = [['Number', 'Net', 'Impedance SPEC (Ω)', 'Layer', 'Signal Type', 'Width/Spacing']]
        # 先单根
        num = 0
        signal_type = 'Single-ended'
        for single_outer_key, single_outer_value in part_outer_single_dict.items():
            single_outer_net_list = single_outer_value
            single_outer_width, single_outer_layer = single_outer_key.split(' ')
            single_outer_impedance = self.outer_ws_impedance_dict[single_outer_width]
            for single_outer_net in single_outer_net_list:
                num += 1
                output_list.append([num, single_outer_net, single_outer_impedance,
                                    single_outer_layer, signal_type, single_outer_width])

        for single_inner_key, single_inner_value in part_inner_single_dict.items():
            single_inner_net_list = single_inner_value
            single_inner_width, single_inner_layer = single_inner_key.split(' ')
            single_inner_impedance = self.inner_ws_impedance_dict[single_inner_width]
            for single_inner_net in single_inner_net_list:
                num += 1
                output_list.append([num, single_inner_net, single_inner_impedance,
                                    single_inner_layer, signal_type, single_inner_width])

        # 再差分
        signal_type = 'Differential'
        for diff_outer_key, diff_outer_value in part_outer_diff_dict.items():
            diff_outer_net_list = diff_outer_value
            diff_outer_width, diff_outer_spacing, diff_outer_layer = diff_outer_key.split(' ')
            diff_outer_impedance = self.outer_ws_impedance_dict[diff_outer_width + ' ' + diff_outer_spacing]
            try:
                diff_outer_width = str(int(float(diff_outer_width)))
            except:
                diff_outer_width = str(float(diff_outer_width))
            try:
                diff_outer_spacing = str(int(float(diff_outer_spacing)))
            except:
                diff_outer_spacing = str(float(diff_outer_spacing))
            for i in range(0, len(diff_outer_net_list), 2):
                num += 1
                output_list.append([num, diff_outer_net_list[i] + '/' + diff_outer_net_list[i + 1],
                                    diff_outer_impedance, diff_outer_layer, signal_type, "'" + diff_outer_width + '/'
                                    + diff_outer_spacing + '/' + diff_outer_width])

        for diff_inner_key, diff_inner_value in part_inner_diff_dict.items():
            diff_inner_net_list = diff_inner_value
            diff_inner_width, diff_inner_spacing, diff_inner_layer = diff_inner_key.split(' ')
            diff_inner_impedance = self.inner_ws_impedance_dict[diff_inner_width + ' ' + diff_inner_spacing]
            try:
                diff_inner_width = str(int(float(diff_inner_width)))
            except:
                diff_inner_width = str(float(diff_inner_width))
            try:
                diff_inner_spacing = str(int(float(diff_inner_spacing)))
            except:
                diff_inner_spacing = str(float(diff_inner_spacing))
            for i in range(0, len(diff_inner_net_list), 2):
                num += 1
                output_list.append([num, diff_inner_net_list[i] + '/' + diff_inner_net_list[i + 1],
                                    diff_inner_impedance, diff_inner_layer, signal_type, "'" + diff_inner_width + '/'
                                    + diff_inner_spacing + '/' + diff_inner_width])

        # 设置excel格式
        first_row_idx = (table_ind[0], table_ind[1] + len(output_list[0]) - 1)
        # print(table_ind)
        # # sht.range(first_row_idx).api.Font.Name = 'Arial'
        # sht.range(table_ind).api.Font.Bold = True
        # sht.range(table_ind).api.Font.Color = '#0000ff'
        column_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
                       'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
                       'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP',
                       'AQ',
                       'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ',
                       'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP',
                       'BQ',
                       'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ',
                       'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP',
                       'CQ',
                       'CR', 'CS', 'CT', 'CU', 'CV', 'CW', 'CX', 'CY', 'CZ',
                       'DA', 'DB', 'DC', 'DD', 'DE', 'DF', 'DG', 'DH', 'DI', 'DJ', 'DK', 'DL', 'DM', 'DN', 'DO', 'DP',
                       'DQ',
                       'DR', 'DS', 'DT', 'DU', 'DV', 'DW', 'DX', 'DY', 'DZ',
                       'EA', 'EB', 'EC', 'ED', 'EE', 'EF', 'EG', 'EH', 'EI', 'EJ', 'EK', 'EL', 'EM', 'EN', 'EO', 'EP',
                       'EQ',
                       'ER', 'ES', 'ET', 'EU', 'EV', 'EW', 'EX', 'EY', 'EZ',
                       'FA', 'FB', 'FC', 'FD', 'FE', 'FF', 'FG', 'FH', 'FI', 'FJ', 'FK', 'FL', 'FM', 'FN', 'FO', 'FP',
                       'FQ',
                       'FR', 'FS', 'FT', 'FU', 'FV', 'FW', 'FX', 'FY', 'FZ',
                       'GA', 'GB', 'GC', 'GD', 'GE', 'GF', 'GG', 'GH', 'GI', 'GJ', 'GK', 'GL', 'GM', 'GN', 'GO', 'GP',
                       'GQ',
                       'GR', 'GS', 'GT', 'GU', 'GV', 'GW', 'GX', 'GY', 'GZ',
                       'HA', 'HB', 'HC', 'HD', 'HE', 'HF', 'HG', 'HH', 'HI', 'HJ', 'HK', 'HL', 'HM', 'HN', 'HO', 'HP',
                       'HQ',
                       'HR', 'HS', 'HT', 'HU', 'HV', 'HW', 'HX', 'HY', 'HZ',
                       'IA', 'IB', 'IC', 'ID', 'IE', 'IF', 'IG', 'IH', 'II', 'IJ', 'IK', 'IL', 'IM', 'IN', 'IO', 'IP',
                       'IQ',
                       'IR', 'IS', 'IT', 'IU', 'IV', 'IW', 'IX', 'IY', 'IZ',
                       'JA', 'JB', 'JC', 'JD', 'JE', 'JF', 'JG', 'JH', 'JI', 'JJ', 'JK', 'JL', 'JM', 'JN', 'JO', 'JP',
                       'JQ',
                       'JR', 'JS', 'JT', 'JU', 'JV', 'JW', 'JX', 'JY', 'JZ',
                       'KA', 'KB', 'KC', 'KD', 'KE', 'KF', 'KG', 'KH', 'KI', 'KJ', 'KK', 'KL', 'KM', 'KN', 'KO', 'KP',
                       'KQ',
                       'KR', 'KS', 'KT', 'KU', 'KV', 'KW', 'KX', 'KY', 'KZ',
                       'LA', 'LB', 'LC', 'LD', 'LE', 'LF', 'LG', 'LH', 'LI', 'LJ', 'LK', 'LL', 'LM', 'LN', 'LO', 'LP',
                       'LQ',
                       'LR', 'LS', 'LT', 'LU', 'LV', 'LW', 'LX', 'LY', 'LZ',
                       'MA', 'MB', 'MC', 'MD', 'ME', 'MF', 'MG', 'MH', 'MI', 'MJ', 'MK', 'ML', 'MM', 'MN', 'MO', 'MP',
                       'MQ',
                       'MR', 'MS', 'MT', 'MU', 'MV', 'MW', 'MX', 'MY', 'MZ',
                       'NA', 'NB', 'NC', 'ND', 'NE', 'NF', 'NG', 'NH', 'NI', 'NJ', 'NK', 'NL', 'NM', 'NN', 'NO', 'NP',
                       'NQ',
                       'NR', 'NS', 'NT', 'NU', 'NV', 'NW', 'NX', 'NY', 'NZ',
                       'OA', 'OB', 'OC', 'OD', 'OE', 'OF', 'OG', 'OH', 'OI', 'OJ', 'OK', 'OL', 'OM', 'ON', 'OO', 'OP',
                       'OQ',
                       'OR', 'OS', 'OT', 'OU', 'OV', 'OW', 'OX', 'OY', 'OZ',
                       'PA', 'PB', 'PC', 'PD', 'PE', 'PF', 'PG', 'PH', 'PI', 'PJ', 'PK', 'PL', 'PM', 'PN', 'PO', 'PP',
                       'PQ',
                       'PR', 'PS', 'PT', 'PU', 'PV', 'PW', 'PX', 'PY', 'PZ',
                       'QA', 'QB', 'QC', 'QD', 'QE', 'QF', 'QG', 'QH', 'QI', 'QJ', 'QK', 'QL', 'QM', 'QN', 'QO', 'QP',
                       'QQ',
                       'QR', 'QS', 'QT', 'QU', 'QV', 'QW', 'QX', 'QY', 'QZ',
                       'RA', 'RB', 'RC', 'RD', 'RE', 'RF', 'RG', 'RH', 'RI', 'RJ', 'RK', 'RL', 'RM', 'RN', 'RO', 'RP',
                       'RQ',
                       'RR', 'RS', 'RT', 'RU', 'RV', 'RW', 'RX', 'RY', 'RZ',
                       'SA', 'SB', 'SC', 'SD', 'SE', 'SF', 'SG', 'SH', 'SI', 'SJ', 'SK', 'SL', 'SM', 'SN', 'SO', 'SP',
                       'SQ',
                       'SR', 'SS', 'ST', 'SU', 'SV', 'SW', 'SX', 'SY', 'SZ',
                       'TA', 'TB', 'TC', 'TD', 'TE', 'TF', 'TG', 'TH', 'TI', 'TJ', 'TK', 'TL', 'TM', 'TN', 'TO', 'TP',
                       'TQ',
                       'TR', 'TS', 'TT', 'TU', 'TV', 'TW', 'TX', 'TY', 'TZ',
                       'UA', 'UB', 'UC', 'UD', 'UE', 'UF', 'UG', 'UH', 'UI', 'UJ', 'UK', 'UL', 'UM', 'UN', 'UO', 'UP',
                       'UQ',
                       'UR', 'US', 'UT', 'UU', 'UV', 'UW', 'UX', 'UY', 'UZ',
                       'VA', 'VB', 'VC', 'VD', 'VE', 'VF', 'VG', 'VH', 'VI', 'VJ', 'VK', 'VL', 'VM', 'VN', 'VO', 'VP',
                       'VQ',
                       'VR', 'VS', 'VT', 'VU', 'VV', 'VW', 'VX', 'VY', 'VZ',
                       'WA', 'WB', 'WC', 'WD', 'WE', 'WF', 'WG', 'WH', 'WI', 'WJ', 'WK', 'WL', 'WM', 'WN', 'WO', 'WP',
                       'WQ',
                       'WR', 'WS', 'WT', 'WU', 'WV', 'WW', 'WX', 'WY', 'WZ',
                       'XA', 'XB', 'XC', 'XD', 'XE', 'XF', 'XG', 'XH', 'XI', 'XJ', 'XK', 'XL', 'XM', 'XN', 'XO', 'XP',
                       'XQ',
                       'XR', 'XS', 'XT', 'XU', 'XV', 'XW', 'XX', 'XY', 'XZ',
                       'YA', 'YB', 'YC', 'YD', 'YE', 'YF', 'YG', 'YH', 'YI', 'YJ', 'YK', 'YL', 'YM', 'YN', 'YO', 'YP',
                       'YQ',
                       'YR', 'YS', 'YT', 'YU', 'YV', 'YW', 'YX', 'YY', 'YZ',
                       'ZA', 'ZB', 'ZC', 'ZD', 'ZE', 'ZF', 'ZG', 'ZH', 'ZI', 'ZJ', 'ZK', 'ZL', 'ZM', 'ZN', 'ZO', 'ZP',
                       'ZQ',
                       'ZR', 'ZS', 'ZT', 'ZU', 'ZV', 'ZW', 'ZX', 'ZY', 'ZZ'
                       ]
        last_idx = (table_ind[0] + len(output_list) - 1, table_ind[1] + len(output_list[0]) - 1)
        # print(len(output_list))
        # print(last_idx)
        # 设置字体行高
        xw.Range('{}{}:{}{}'.format(column_list[table_ind[1] - 1], table_ind[0],
                                    column_list[last_idx[1] - 1], last_idx[0])).row_height = 30
        xw.Range('{}{}:{}{}'.format(column_list[table_ind[1] - 1], table_ind[0],
                                    column_list[last_idx[1] - 1], last_idx[0])).api.Font.Name = 'Arial'
        # 设置垂直水平居中
        xw.Range('{}{}:{}{}'.format(column_list[table_ind[1] - 1], table_ind[0],
                                    column_list[last_idx[1] - 1], last_idx[0])).api.VerticalAlignment = -4108
        xw.Range('{}{}:{}{}'.format(column_list[table_ind[1] - 1], table_ind[0],
                                    column_list[last_idx[1] - 1], last_idx[0])).api.HorizontalAlignment = -4108
        # 设置列自适应
        sht.range(table_ind).value = output_list
        xw.Range('{}{}:{}{}'.format(column_list[table_ind[1] - 1], table_ind[0],
                                    column_list[first_row_idx[1] - 1], first_row_idx[0])).api.Font.Color = 0xff0000
        xw.Range('{}{}:{}{}'.format(column_list[table_ind[1] - 1], table_ind[0],
                                    column_list[first_row_idx[1] - 1], first_row_idx[0])).api.Font.Bold = True
        xw.Range('{}{}:{}{}'.format(column_list[table_ind[1] - 1], table_ind[0],
                                    column_list[last_idx[1] - 1], last_idx[0])).api.Borders.LineStyle = 1
        xw.Range('{}{}:{}{}'.format(column_list[table_ind[1] - 1], table_ind[0],
                                    column_list[last_idx[1] - 1], last_idx[0])).columns.autofit()
        # 设置行高
        # sht.range(range_idx).row_height = 30
        wb.save()
        wb.close()
        app.quit()

        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(self.output_path)

        try:
            sht2 = wb.sheets['detail net info']
            sht2.clear()
        except:
            wb.sheets.add('detail net info')
            sht2 = wb.sheets['detail net info']
            sht2.clear()

        # 将详细信息填入'detail net info' sheet中
        # 先单根
        sht2.range('A1').value = 'Single-Ended'
        xw.Range('A1').api.Font.Bold = True
        xw.Range('A1').api.Font.size = 16
        xw.Range('A1:AA1000').api.Font.Name = 'Arial'

        row_idx = 1
        for key, value in self.outer_single_width_net_dict.items():
            sht2.range((2, row_idx)).value = key
            sht2.range((2, row_idx)).api.Font.Bold = True
            input_value = [[i] for i in value]
            sht2.range((3, row_idx)).value = input_value
            row_idx += 1

        for key, value in self.inner_single_width_net_dict.items():
            sht2.range((2, row_idx)).value = key
            sht2.range((2, row_idx)).api.Font.Bold = True
            input_value = [[i] for i in value]
            sht2.range((3, row_idx)).value = input_value
            row_idx += 1

        # 后差分
        row_idx += 1
        sht2.range((1, row_idx)).value = 'Differential'
        xw.Range((1, row_idx)).api.Font.size = 16
        xw.Range((1, row_idx)).api.Font.Bold = True
        for key, value in self.outer_diff_width_net_dict.items():
            sht2.range((2, row_idx)).value = key
            sht2.range((2, row_idx)).api.Font.Bold = True
            input_value = [[i] for i in value]
            sht2.range((3, row_idx)).value = input_value
            row_idx += 1

        for key, value in self.inner_diff_width_net_dict.items():
            sht2.range((2, row_idx)).value = key
            sht2.range((2, row_idx)).api.Font.Bold = True
            input_value = [[i] for i in value]
            sht2.range((3, row_idx)).value = input_value
            row_idx += 1

        sht2.autofit()

        wb.save()
        wb.close()
        app.quit()


choose_test_point = ChooseTestPoint()
choose_test_point.get_net_list_from_checklist()
choose_test_point.get_all_specifications_from_output_file()
choose_test_point.export_brd_report()
choose_test_point.create_output_file()
