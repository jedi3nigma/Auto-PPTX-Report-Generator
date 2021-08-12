import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import six
import numpy as np

sns.set_style(style='white')


class Chart():
    def __init__(self, row_dim, col_dim, *paths, **kwargs):
        self.paths_count = len(paths)
        if self.paths_count > 1:
            for i, path in enumerate(paths):
                super().__setattr__('path_{}'.format(i), path)
        elif self.paths_count == 1:
            self.path = paths[0]
        else:
            self.path = None
        self.row_dim = row_dim
        self.col_dim = col_dim

    @staticmethod
    def _plot_setup(row_dim, col_dim, x_lab, y_lab, b_rt_spine, b_lt_spine, b_tp_spine):
        fig, ax = plt.subplots(figsize=(row_dim, col_dim))
        ax.spines['left'].set_visible(b_lt_spine)
        ax.spines['right'].set_visible(b_rt_spine)
        ax.spines['top'].set_visible(b_tp_spine)
        ax.set(xlabel=x_lab, ylabel=y_lab)
        return fig, ax

    def _save(self, **kwargs):
        bbox_inches = kwargs.get('bbox_inches', 'tight')
        pad_inches = kwargs.get('pad_inches', 0.1)
        if self.paths_count > 1:
            for i in range(self.paths_count):
                file_path = getattr(self, 'path_{}'.format(i))
                plt.savefig(file_path, bbox_inches=bbox_inches,
                            pad_inches=pad_inches)
        elif self.paths_count == 1:
            plt.savefig(self.path, bbox_inches=bbox_inches,
                        pad_inches=pad_inches)
        else:
            pass


class Box(Chart):
    def __init__(self, row_dim, col_dim, data, *paths, **kwargs):
        super().__init__(row_dim, col_dim, *paths, **kwargs)
        self.data = data
        self.combo_chart = kwargs.get('combo', False)

    def plot(self, x, y, **kwargs):
        x_lab = kwargs.get('x_lab', '')
        y_lab = kwargs.get('y_lab', '')
        b_rt_spine = kwargs.get('b_rt_spine', False)
        b_lt_spine = kwargs.get('b_lt_spine', True)
        b_tp_spine = kwargs.get('b_tp_spine', False)
        linewidth = kwargs.get('linewidth', 0.5)
        palette = kwargs.get('palette', 'viridis')
        orient = kwargs.get('orient', 'h')
        color = kwargs.get('color', 'None')

        fig, ax = self._plot_setup(
            self.row_dim, self.col_dim, x_lab, y_lab, b_rt_spine, b_lt_spine, b_tp_spine)
        ch = sns.boxplot(y=y, x=x, data=self.data, color=color,
                         palette=palette, orient=orient, linewidth=linewidth, ax=ax)
        ax.set(xlabel=x_lab, ylabel=y_lab)

    def save(self, **kwargs):
        super()._save(**kwargs)


class Line(Chart):
    def __init__(self, row_dim, col_dim, data, *paths, **kwargs):
        super().__init__(row_dim, col_dim, *paths, **kwargs)
        self.multi_line = kwargs.get('multi_line', False)
        self.idx_data = kwargs.get('idx_data', None)
        self.combo_chart = kwargs.get('combo', False)
        self.main_data = data

    def plot(self, x, y, multi_line_fmt=1, **kwargs):
        x_lab = kwargs.get('x_lab', '')
        y_lab = kwargs.get('y_lab', '')
        combo_ax = kwargs.get('combo_ax', None)
        b_rt_spine = kwargs.get('b_rt_spine', False)
        b_lt_spine = kwargs.get('b_lt_spine', True)
        b_tp_spine = kwargs.get('b_tp_spine', False)
        linewidth = kwargs.get('linewidth', 0.5)
        palette = kwargs.get('palette', 'muted')
        alpha = kwargs.get('alpha', 1)
        color = kwargs.get('color', 'red')
        selector_col = kwargs.get('selector_col', None)
        x_axis_time = kwargs.get('x_axis_time', False)

        if not self.combo_chart:
            fig, ax = self._plot_setup(
                self.row_dim, self.col_dim, x_lab, y_lab, b_rt_spine, b_lt_spine, b_tp_spine)

        if self.multi_line:
            if multi_line_fmt == 1:
                selector_alpha_map = dict(
                    zip(self.idx_data, np.arange(0.3, 1, step=1/len(self.idx_data))))
                for selector, a in selector_alpha_map.items():
                    filt_data = self.main_data[self.main_data[selector_col] == selector]
                    ch = sns.lineplot(x=x, y=y, data=filt_data,
                                      palette=palette, alpha=a, ax=ax)
                ch.legend(labels=selector_alpha_map.keys())
            elif multi_line_fmt == 2:
                colors = kwargs.get(
                    'm_colors', list(mcolors.TABLEAU_COLORS.values()))
                y_lab_list = kwargs.get('y_list', None)
                leg_lab_list = kwargs.get('leg_lab_list', None)
                colors = colors[:len(y_lab_list)]
                if (y_lab_list is None) | (leg_lab_list is None):
                    print('Error: please provide y-labels and/or legend labels')
                    return

                try:
                    assert (isinstance(y_lab_list, list)) & (len(
                        colors) == len(y_lab_list) == len(leg_lab_list))
                except AssertionError as msg:
                    print('y variable is not a list or length of y')
                    return

                for var, color, label in zip(y_lab_list, colors, leg_lab_list):
                    ch = sns.lineplot(x=x, y=var, data=self.main_data,
                                      color=color, label=label, ax=ax)
                ax.legend()
        else:
            if (self.combo_chart) & (combo_ax != None):
                ch = sns.lineplot(x=x, y=y, data=self.main_data, color=color,
                                  palette=palette, alpha=alpha, ax=combo_ax)
            else:
                ch = sns.lineplot(x=x, y=y, data=self.main_data, color=color,
                                  palette=palette, alpha=alpha, ax=ax)

        if x_axis_time:
            plt.gcf().autofmt_xdate()

        if (not self.combo_chart) | (combo_ax != None):
            self.save()
            return ch
        elif self.combo_chart:
            combo_ax = plt.twinx()
            # combo_ax.set(ylabel=y_lab)
            # combo_ax.spines['top'].set_visible(False)
            return ch, combo_ax

    def save(self, **kwargs):
        super()._save(**kwargs)


class Bar(Chart):
    def __init__(self, row_dim, col_dim, data, *paths, **kwargs):
        super().__init__(row_dim, col_dim, *paths, **kwargs)
        self.data = data
        self.combo_chart = kwargs.get('combo', False)

    def plot(self, x, y, **kwargs):
        x_lab = kwargs.get('x_lab', '')
        y_lab = kwargs.get('y_lab', '')
        b_rt_spine = kwargs.get('b_rt_spine', False)
        b_lt_spine = kwargs.get('b_lt_spine', True)
        b_tp_spine = kwargs.get('b_tp_spine', False)
        linewidth = kwargs.get('linewidth', 0.5)
        palette = kwargs.get('palette', 'viridis')
        orient = kwargs.get('orient', 'h')
        color = kwargs.get('color', 'None')

        fig, ax = self._plot_setup(
            self.row_dim, self.col_dim, x_lab, y_lab, b_rt_spine, b_lt_spine, b_tp_spine)

        ch = sns.barplot(y=y, x=x, data=self.data, color=color,
                         orient=orient, linewidth=linewidth, ax=ax)
        ax.set(xlabel=x_lab, ylabel=y_lab)

        if not self.combo_chart:
            self.save()
            return ch
        else:
            combo_ax = plt.twinx()
            return ch, combo_ax

    def save(self, **kwargs):
        super()._save(**kwargs)


class Table():
    def __init__(self, row_height, col_width, data, path, **kwargs):
        self.row_dim = row_height
        self.col_dim = col_width
        self.data = data
        self.path = path

    def plot(self, **kwargs):
        font_size = kwargs.get('font_size', 12)
        ch = self._render_mpl_table(
            self.data, col_width=self.col_dim, font_size=font_size, row_height=self.row_dim, **kwargs)
        self._save(ch, self.path)

    def _render_mpl_table(self, data, col_width=3.0, row_height=0.625, font_size=14, header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w', bbox=[0, 0, 1, 1], header_columns=0, ax=None, **kwargs):
        if ax is None:
            size = (np.array(data.shape[::-1]) + np.array([0, 1])
                    ) * np.array([col_width, row_height])
            fig, ax = plt.subplots(figsize=size)
            ax.axis('off')

        mpl_table = ax.table(cellText=data.values, bbox=bbox,
                             colLabels=data.columns, **kwargs)

        mpl_table.auto_set_font_size(False)
        mpl_table.set_fontsize(font_size)

        for k, cell in six.iteritems(mpl_table._cells):
            cell.set_edgecolor(edge_color)
            if k[0] == 0 or k[1] < header_columns:
                cell.set_text_props(weight='bold', color='w')
                cell.set_facecolor(header_color)
            else:
                cell.set_facecolor(row_colors[k[0] % len(row_colors)])
        return ax

    @staticmethod
    def _save(ax, path, **kwargs):
        bbox_inches = kwargs.get('bbox_inches', 'tight')
        pad_inches = kwargs.get('pad_inches', 0.1)
        ax.get_figure().savefig(path, bbox_inches=bbox_inches,
                                pad_inches=pad_inches)
