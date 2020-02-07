# -*- coding: utf-8 -*-
"""
Created on Wed May 15 22:51:04 2019

@author: v.shkaberda
"""
from _version import __version__
from checkboxtreeview import CheckboxTreeview
from calendar import month_name
from datetime import date, datetime
from decimal import Decimal
from label_grid import LabelGrid
from multiselect import MultiselectMenu
from tkcalendar import DateEntry
from tkinter import ttk, messagebox
import tkinter.font as tkFont
from tkHyperlinkManager import HyperlinkManager
from math import floor
from xl import export_to_excel
import locale
import os, zlib
import tkinter as tk

# example of subsription and default recipient
EMAIL_TO = b'\xd0\xa4\xd0\xbe\xd0\xb7\xd0\xb7\xd0\xb8|\
\xd0\x9b\xd0\xbe\xd0\xb3\xd0\xb8\xd1\x81\xd1\x82\xd0\xb8\xd0\xba\xd0\xb0|\
\xd0\x90\xd0\xbd\xd0\xb0\xd0\xbb\xd0\xb8\xd1\x82\xd0\xb8\xd0\xba\xd0\xb8'.decode()
# example of path to independent report
REPORT_PATH = zlib.decompress(b'x\x9c\x8b\x89I\xcb\xaf\xaa\xaa\xd4\xcbI\xcc\
\x8bq\xc9O.\xcdM\xcd+)\x8e\xf1\xc8\xcfI\xc9\xccK\x8fqI-H,*\x81\x88\xf9\xe4\
\xa7g\x16\x97df\'\xc6\xb8e\xe6\xc5\\Xpa\xc3\xc5\xc6\x0b\xfb/6\\\xd8za\x0b\x10\
\xef\x06\xe2\xbd\x17v\\\xd8\x1a\x7fa;P\xaa\t(\x01$c.L\xb9\xb0\xef\xc2~\x85\x0b\
\xfb\x80"\xed\x17\xb6\x02\xc9n\x00\x9b\x8c?\xef').decode()


class PaymentsError(Exception):
    """Base class for exceptions in this module."""
    pass


class IncorrectFloatError(PaymentsError):
    """ Exception raised if sum is not converted to float.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """
    def __init__(self, expression, message='Введена некорректная сумма'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class MonthFilterError(PaymentsError):
    """ Exception raised if month don't chosen in filter.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """
    def __init__(self, expression, message='Не выбран месяц'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class MonthChangedError(PaymentsError):
    """ Exception raised if month is changed when altering payment.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """
    def __init__(self, expression, message='Запрещено менять месяц'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class NoRightsToFillCreateFormError(PaymentsError):
    """ Exception raised when trying to fill CreateForm using data with
    no rights to be used for creation.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """
    def __init__(self, expression,
                 message='Нет прав для использования данного МВЗ/офиса'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class PeriodExceededError(PaymentsError):
    """ Exception raised if period is exceeded when altering payment.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """
    def __init__(self, expression, message='Превышен допустимый период'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class SumExceededError(PaymentsError):
    """ Exception raised if sum is exceeded when altering payment.

    Attributes:
        expression - input expression in which the error occurred;
        message - explanation of the error.
    """
    def __init__(self, expression, message='Превышена первичная сумма'):
        self.expression = expression
        self.message = message
        super().__init__(self.expression, self.message)


class AccessError(tk.Tk):
    """ Raise an error when user doesn't have permission to work with app.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
            'Ошибка доступа',
            'Нет прав для работы с приложением.\n'
            'Для получения прав обратитесь на рассылку ' + EMAIL_TO
        )
        self.destroy()


class LoginError(tk.Tk):
    """ Raise an error when user doesn't have permission to work with db.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
                'Ошибка подключения',
                'Нет прав для работы с сервером.\n'
                'Обратитесь на рассылку ' + EMAIL_TO
        )
        self.destroy()


class NetworkError(tk.Tk):
    """ Raise a message about network error.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showerror(
                'Ошибка cети',
                'Возникла общая ошибка сети.\nПерезапустите приложение'
        )
        self.destroy()


class RestartRequiredAfterUpdateError(tk.Tk):
    """ Raise a message about restart needed after update.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showinfo(
                'Необходима перезагрузка',
                'Выполнено критическое обновление.\nПерезапустите приложение'
        )
        self.destroy()


class UnexpectedError(tk.Tk):
    """ Raise a message when an unexpected exception occurs.
    """
    def __init__(self, *args):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showinfo(
                'Непредвиденное исключение',
                'Возникло непредвиденное исключение\n' + '\n'.join(map(str, args))
        )
        self.destroy()


class StringSumVar(tk.StringVar):
    """ Contains function that returns var formatted in a such way, that
        it can be converted into a float without an error.
    """
    def get_float_form(self, *args, **kwargs):
        return super().get(*args, **kwargs).replace(' ', '').replace(',', '.')


class PaymentApp(tk.Tk):
    def __init__(self, **kwargs):
        super().__init__()
        self.title('Заявки на оплату')
        self.iconbitmap('resources/payment.ico')
        # store the state of PreviewForm
        self.state_PreviewForm = 'normal'
        # geometry_storage {Framename:(width, height)}
        self._geometry = {'PreviewForm': (1200, 600),
                          'CreateForm': (750, 440)}
        # Virtual event for creating request
        self.event_add("<<create>>", "<Control-S>", "<Control-s>",
                       "<Control-Ucircumflex>", "<Control-ucircumflex>",
                       "<Control-twosuperior>", "<Control-threesuperior>",
                       "<KeyPress-F5>")
        self.bind_all("<Key>", self._onKeyRelease, '+')
        self.bind("<<create>>", self._create_request)
        self.active_frame = None
        # handle the window close event
        self.protocol("WM_DELETE_WINDOW", self._quit)
        # hide until all frames have been created
        self.withdraw()
        # To import months names in cyrillic
        locale.setlocale(locale.LC_ALL, 'RU')
        # Customize header style (used in PreviewForm)
        style = ttk.Style()
        try:
            style.element_create("HeaderStyle.Treeheading.border", "from", "default")
            style.layout("HeaderStyle.Treeview.Heading", [
                ("HeaderStyle.Treeheading.cell", {'sticky': 'nswe'}),
                ("HeaderStyle.Treeheading.border", {'sticky': 'nswe', 'children': [
                    ("HeaderStyle.Treeheading.padding", {'sticky': 'nswe', 'children': [
                        ("HeaderStyle.Treeheading.image", {'side': 'right', 'sticky':''}),
                        ("HeaderStyle.Treeheading.text", {'sticky': 'we'})
                    ]})
                ]}),
            ])
            style.configure("HeaderStyle.Treeview.Heading",
                background="#eaeaea", foreground="black", relief='groove', font=('Arial', 8))
            style.map("HeaderStyle.Treeview.Heading",
                      relief=[('active', 'sunken'), ('pressed', 'flat')])

            style.map('ButtonGreen.TButton')
            style.configure('ButtonGreen.TButton', foreground='green')

            style.map('ButtonRed.TButton')
            style.configure('ButtonRed.TButton', foreground='red')

            style.configure("TMenubutton", background='white')
        except tk.TclError:
            # if during debug previous tk wasn't destroyed
            # and style remains modified
            pass

        # the container is where we'll stack a bunch of frames
        # the one we want to be visible will be raised above others
        container = tk.Frame(self)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self._frames = {}
        for F in (PreviewForm, CreateForm):
            frame_name = F.__name__
            frame = F(parent=container, controller=self, **kwargs)
            self._frames[frame_name] = frame
            # put all of them in the same location
            frame.grid(row=0, column=0, sticky='nsew')

        self._show_frame('PreviewForm')
        # restore after withdraw
        self.deiconify()

    def _center_window(self, w, h):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        start_x = int((screen_width/2) - (w/2))
        start_y = int((screen_height/2) - (h*0.55))

        self.geometry('{}x{}+{}+{}'.format(w, h, start_x, start_y))

    def _fill_CreateForm(self, МВЗ, Офис, Категория, Контрагент, Описание,
                         **kwargs):
        """ Control function to transfer data from Preview- to CreateForm. """
        frame = self._frames['CreateForm']
        frame._fill_from_PreviewForm(МВЗ, Офис, Категория, Контрагент, Описание)

    def _onKeyRelease(*args):
        event = args[1]
        # check if Ctrl pressed
        if not (event.state == 12 or event.state == 14):
            return
        if event.keycode == 88 and event.keysym.lower() != 'x':
            event.widget.event_generate("<<Cut>>")
        elif event.keycode == 86 and event.keysym.lower() != 'v':
            event.widget.event_generate("<<Paste>>")
        elif event.keycode == 67 and event.keysym.lower() != 'c':
            event.widget.event_generate("<<Copy>>")

    def _show_frame(self, frame_name):
        """ Show a frame for the given frame name. """
        if frame_name == 'CreateForm':
            # since we have only two forms, when we activating CreateForm
            # we know by exception that PreviewForm is active
            self.state_PreviewForm = self.state()
            self.state('normal')
            self.resizable(width=False, height=False)
        else:
            self.state(self.state_PreviewForm)
            self.resizable(width=True, height=True)
        frame = self._frames[frame_name]
        frame.tkraise()
        self._center_window(*(self._geometry[frame_name]))
        # Make sure active_frame changes in case of network error
        try:
            if frame_name in ('PreviewForm'):
                frame._resize_columns()
                frame._refresh()
        finally:
            self.active_frame = frame_name

    def _create_request(self, event):
        """
        Creates request when hotkey is pressed if active_frame is CreateForm.
        """
        if self.active_frame == 'CreateForm':
            self._frames[self.active_frame]._create_request()

    def _quit(self):
        if self.active_frame != 'PreviewForm':
            self._show_frame('PreviewForm')
        elif messagebox.askokcancel("Выход", "Выйти из приложения?"):
            self.destroy()


class PaymentFrame(tk.Frame):
    def __init__(self, parent, controller, connection, user_info,
                 mvz, categories, pay_conditions):
        super().__init__(parent)
        self.parent = parent
        self.controller = controller
        self.conn = connection
        self.categories = dict(categories)
        self.pay_conditions = dict(pay_conditions)
        # {mvzSAP: [mvzname, [office1, office2, ...]], ...}
        self.mvz = {}
        if isinstance(self, PreviewForm):
            self.mvz['Все'] = (None, ('Все',))
        for mvzSAP, mvzname, office in mvz:
            try:
                self.mvz[mvzname][1].append(office)
            except KeyError:
                self.mvz[mvzname] = [mvzSAP, [office]]
        self.user_info = user_info
        # Often used info
        self.userID = user_info.UserID
        self.PayConditionsDefaultID = user_info.PayConditionsID


    def _add_user_label(self, parent):
        """ Adds user name in top right corner. """
        user_label = tk.Label(parent, text='Пользователь: ' +
                     self.user_info.ShortUserName + '  Версия ' + __version__,
                     font=('Arial', 8))
        user_label.pack(side=tk.RIGHT, anchor=tk.NE)

    def _format_float(self, sum_float):
        return '{:,.2f}'.format(sum_float).replace(',', ' ').replace('.', ',')

    def _on_focus_in_format_sum(self, event):
        """ Convert str into float in binded to entry variable when focus in.
        """
        varname = str(event.widget.cget("textvariable"))
        sum_entry = event.widget.get().replace(' ', '')
        event.widget.setvar(varname, sum_entry)

    def _on_focus_out_format_sum(self, event):
        """ Convert float into str in binded to entry variable when focus out.
        """
        if not event.widget.get().replace(',', '.'):
            return
        sum_entry = float(event.widget.get().replace(',', '.'))
        varname = str(event.widget.cget("textvariable"))
        event.widget.setvar(varname, self._format_float(sum_entry))

    def _validate_sum(self, sum_entry):
        """ Validation of entries that contains sum. """
        sum_entry = sum_entry.replace(',', '.').replace(' ', '')
        try:
            if not sum_entry or 0 <= float(sum_entry) < 10**9:
                return True
        except (TypeError, ValueError):
            return False
        return False

    def get_mvzSAP(self, mvz):
        return self.mvz[mvz][0]

    def get_offices(self, mvz):
        return self.mvz[mvz][1]


class CreateForm(PaymentFrame):
    def __init__(self, parent, controller, connection, user_info,
                 mvz, categories, pay_conditions, approvals_for_first_stage, **kwargs):
        super().__init__(parent, controller, connection, user_info,
                         mvz, categories, pay_conditions)
        self.approvals_for_first_stage = dict(approvals_for_first_stage)
        # Top Frame with description and user name
        top = tk.Frame(self, name='top_cf', padx=5)
        self.main_label = tk.Label(top, text='Форма создания заявки на согласование',
                              padx=10, font=('Arial', 8, 'bold'))

        self.limit_label = tk.Label(top, text='Оставшийся лимит     —',
                              padx=10, font=('Arial', 8, 'bold'), fg='#003db9')
        self.limit_month = tk.Label(top, text='', font=('Arial', 9))
        self.limit_sum = tk.Label(top, text='', font=('Arial', 8, 'bold'))

        self.pay_conditions = dict(pay_conditions)
        # Choose default pay conditions for current user
        for key, value in self.pay_conditions.items():
            if value == self.PayConditionsDefaultID:
                self.PayConditionsDefault = key

        self._add_user_label(top)
        self._top_pack()

        # First Fill Frame with (MVZ, office)
        row1_cf = tk.Frame(self, name='row1_cf', padx=15)

        self.mvz_label = tk.Label(row1_cf, text='МВЗ', padx=10)
        self.mvz_current = tk.StringVar()
        # self.mvz_current.set(self.mvznames[0]) # default value
        self.mvz_box = ttk.OptionMenu(row1_cf, self.mvz_current, '', *self.mvz.keys(),
                                      command=self._restraint_by_mvz)
        self.mvz_box.config(width=40)
        self.mvz_sap = tk.Label(row1_cf, padx=6, bg='lightgray', width=11)

        self.category_label = tk.Label(row1_cf, text='Категория', padx=10)
        self.category_box = ttk.Combobox(row1_cf, width=29, state='readonly')
        self.category_box['values'] = list(self.categories)

        self._row1_pack()

        # Second Fill Frame with (contragent, CSP)
        row2_cf = tk.Frame(self, name='row2_cf', padx=15)

        self.office_label = tk.Label(row2_cf, text='Офис', padx=10)
        self.office_box = ttk.Combobox(row2_cf, width=20, state='disabled')
        self.contragent_label = tk.Label(row2_cf, text='Контрагент')
        self.contragent_entry = tk.Entry(row2_cf, width=28)
        self.csp = tk.Label(row2_cf, text='CSP', padx=10)
        self.csp_entry = tk.Entry(row2_cf, width=26)

        self._row2_pack()

        # Second Fill Frame with (Plan date, Sum, Tax)
        row3_cf = tk.Frame(self, name='row3_cf', padx=15)

        self.plan_date_label = tk.Label(row3_cf, text='Плановая дата', padx=10)
        self.plan_date = tk.StringVar()
        self.plan_date.trace("w", self._check_limit)
        self.plan_date_entry = DateEntry(row3_cf, width=12, state='readonly',
                                         textvariable=self.plan_date, font=('Arial', 9),
                                         selectmode='day', borderwidth=2,
                                         locale='ru_RU')
        self.sum_label = tk.Label(row3_cf, text='Сумма без НДС, грн')
        self.sumtotal = StringSumVar()
        self.sumtotal.set('0,00')
        vcmd = (self.register(self._validate_sum))
        self.sum_entry = tk.Entry(row3_cf, name='sum_entry', width=16,
                                  textvariable=self.sumtotal, validate='all',
                                  validatecommand=(vcmd, '%P')
                                  )
        self.sum_entry.bind("<FocusIn>", self._on_focus_in_format_sum)
        self.sum_entry.bind("<FocusOut>", self._on_focus_out_format_sum)
        self.nds_label = tk.Label(row3_cf, text='НДС', padx=20)
        self.nds = tk.IntVar()
        self.nds.set(20)
        self.nds20 = ttk.Radiobutton(row3_cf, text="20 %", variable=self.nds, value=20)
        self.nds7 = ttk.Radiobutton(row3_cf, text="7 %", variable=self.nds, value=7)
        self.nds0 = ttk.Radiobutton(row3_cf, text="0 %", variable=self.nds, value=0)

        self._row3_pack()

        # Fourth Fill Frame with (cashless)
        row4_cf = tk.Frame(self, name='row4_cf', padx=15)

        self.cashless_label = tk.Label(row4_cf, text='Вид платежа', padx=10)
        self.cashless = ttk.Combobox(row4_cf, width=14, state='readonly')
        self.cashless['values'] = ['наличный','безналичный']
        self.cashless.current(1)
        # Pay conditions variances with fill value as default for user
        self.pay_conditions_label = tk.Label(row4_cf, text='Условия оплаты', padx=10)
        self.pay_conditions_box = ttk.Combobox(row4_cf, width=15,
                                               state='readonly')
        self.pay_conditions_box['values'] = list(self.pay_conditions)
        self.pay_conditions_box.current(list(self.pay_conditions).
                                        index(self.PayConditionsDefault))
        self.initiator_label = tk.Label(row4_cf, text='Инициатор')
        self.initiator_label_name = tk.Entry(row4_cf, width=20)
        sc = tk.Entry(row4_cf, width=15)

        # Text Frame
        text_cf = ttk.LabelFrame(self, text=' Описание заявки ', name='text_cf')

        self.customFont = tkFont.Font(family="Arial", size=10)
        self.desc_text = tk.Text(text_cf, font=self.customFont)  # input and output box
        self.desc_text.configure(width=100)
        self.desc_text.pack(in_=text_cf, expand=True)

        # Approval choose Frame
        appr_cf = tk.Frame(self, name='appr_cf', padx=15)
        self.approval_label = tk.Label(appr_cf, text='Первый этап утверждения:', padx=10)
        self.approval_box = ttk.Combobox(appr_cf, width=40, state='readonly')
        self.approval_box['values'] = list(self.approvals_for_first_stage)

        self._row4_pack()

        # Bottom Frame with buttons
        bottom_cf = tk.Frame(self, name='bottom_cf')

        bt3 = ttk.Button(bottom_cf, text="Назад", width=10,
                         command=lambda: controller._show_frame('PreviewForm'))
        bt3.pack(side=tk.RIGHT, padx=15, pady=10)

        bt2 = ttk.Button(bottom_cf, text="Очистить", width=10,
                         command=self._clear, style='ButtonRed.TButton')
        bt2.pack(side=tk.RIGHT, padx=15, pady=10)

        bt1 = ttk.Button(bottom_cf, text="Создать", width=10,
                         command=self._create_request, style='ButtonGreen.TButton')
        bt1.pack(side=tk.RIGHT, padx=15, pady=10)

        # Pack frames
        top.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        bottom_cf.pack(side=tk.BOTTOM, fill=tk.X)
        appr_cf.pack(side=tk.BOTTOM, fill=tk.X, expand=True, pady=5)
        row1_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row2_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row3_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        row4_cf.pack(side=tk.TOP, fill=tk.X, pady=5)
        text_cf.pack(side=tk.TOP, fill=tk.X, expand=True, padx=15, pady=15)

    def _check_limit(self, *args, **kwargs):
        """ Show remaining limit for month corresponding to month of plan_date.
        """
        plan_date = self.plan_date.get()
        if not plan_date:
            return
        limit = self.conn.get_limit_for_month_by_date(self.userID, self._convert_date(plan_date))
        self.limit_month.configure(text=self._convert_date(plan_date, output="%B %Y") + ': ')
        self.limit_sum.configure(text=self._format_float(limit) + ' грн.')
        self.limit_sum.configure(fg=('black' if limit else 'red'))

    def _clear(self):
        self.mvz_current.set('')
        self.mvz_sap.config(text='')
        self.category_box.set('')
        self.pay_conditions_box.set('')
        self.office_box.set('')
        self.office_box.configure(state="disabled")
        self.contragent_entry.delete(0, tk.END)
        self.csp_entry.delete(0, tk.END)
        self.sumtotal.set('0,00')
        self.nds.set(20)
        self.desc_text.delete("1.0", tk.END)
        self.plan_date_entry.set_date(datetime.now())
        self.approval_box.set('')
        self.cashless.set('безналичный')

    def _convert_date(self, date, output=None):
        """ Take date and convert it into output format.
            If output is None datetime object is returned.

            date: str in format '%d[./]%m[./]%y' or '%d[./]%m[./]%Y'.
            output: str or None, output format.
        """
        date = date.replace('/', '.')
        try:
            dat = datetime.strptime(date, '%d.%m.%y')
        except ValueError:
            dat = datetime.strptime(date, '%d.%m.%Y')
        if output:
            return dat.strftime(output)
        return dat


    def _create_request(self):
        messagetitle = 'Создание заявки'
        sumtotal = float(self.sumtotal.get_float_form()
                         if self.sum_entry.get() else 0)
        is_validated = self._validate_request_creation(messagetitle, sumtotal)
        if not is_validated:
            return

        is_cashless = 1 if self.cashless.get() == 'безналичный' else 0
        first_approval = (self.approvals_for_first_stage[self.approval_box.get()]
                          if self.approval_box.get() else None)
        request = {'mvz': self.mvz_sap.cget('text') or None,
                   'office': self.office_box.get(),
                   'categoryID': self.categories[self.category_box.get()],
                   'contragent': self.contragent_entry.get().strip().replace(
                       '\n','') or None,
                   'csp': self.csp_entry.get() or None,
                   'plan_date': self._convert_date(self.plan_date_entry.get()),
                   'sumtotal': sumtotal,
                   'nds':  self.nds.get(),
                   'text': self.desc_text.get("1.0", tk.END).strip(),
                   'approval': first_approval,
                   'is_cashless': is_cashless,
                   'payconditionsID': self.pay_conditions[self.pay_conditions_box.get()]
                   }
        created_success = self.conn.create_request(userID=self.userID, **request)
        if created_success == 1:
            messagebox.showinfo(
                    messagetitle, 'Заявка создана'
            )
            self._clear()
            self.controller._show_frame('PreviewForm')
        elif created_success == 0:
            dat = self._convert_date(self.plan_date_entry.get(), output="%B %Y")
            messagebox.showerror(
                    messagetitle,
                    'Превышен лимит суммы на {}.\n'
                    'Для повышения лимита обратитесь в отдел контроллинга'
                    .format(dat)
            )
        else:
            messagebox.showerror(
                    messagetitle, 'Произошла ошибка при создании заявки'
            )

    def _fill_from_PreviewForm(self, mvz, office, category, contragent,
                               description):
        """ When button "Создать из заявки" from PreviewForm is activated,
        fill some fields taken from choosed in PreviewForm request.
        """
        self._clear()
        self.mvz_current.set(mvz)
        self.mvz_sap.config(text=self.get_mvzSAP(self.mvz_current.get()) or '')
        self.office_box.set(office)
        self.category_box.set(category)
        self.contragent_entry.insert(0, contragent)
        self.desc_text.insert('end', description.strip() if type(description) == str
                              else description)

    def _restraint_by_mvz(self, event):
        """ Shows mvz_sap that corresponds to chosen MVZ and restraint offices.
            If 1 office is available, choose it, otherwise make box active.
        """
        # tcl language has no notion of None or a null value, so use '' instead
        self.mvz_sap.config(text=self.get_mvzSAP(self.mvz_current.get()) or '')
        offices = self.get_offices(self.mvz_current.get())
        if len(offices) == 1:
            self.office_box.set(offices[0])
            self.office_box.configure(state="disabled")
        else:
            self.office_box.set('')
            self.office_box['values'] = offices
            self.office_box.configure(state="readonly")

    def _row1_pack(self):
        self.mvz_label.pack(side=tk.LEFT)
        self.mvz_box.pack(side=tk.LEFT, padx=5)
        self.mvz_sap.pack(side=tk.LEFT)
        self.category_box.pack(side=tk.RIGHT, padx=5)
        self.category_label.pack(side=tk.RIGHT)

    def _row2_pack(self):
        self.office_label.pack(side=tk.LEFT)
        self.office_box.pack(side=tk.LEFT)
        self.csp_entry.pack(side=tk.RIGHT, padx=5)
        self.csp.pack(side=tk.RIGHT)
        self.contragent_entry.pack(side=tk.RIGHT, padx=5)
        self.contragent_label.pack(side=tk.RIGHT)

    def _row3_pack(self):
        self.plan_date_label.pack(side=tk.LEFT)
        self.plan_date_entry.pack(side=tk.LEFT, padx=5)
        self.nds0.pack(side=tk.RIGHT, padx=7)
        self.nds7.pack(side=tk.RIGHT, padx=7)
        self.nds20.pack(side=tk.RIGHT, padx=8)
        self.nds_label.pack(side=tk.RIGHT)
        self.sum_entry.pack(side=tk.RIGHT, padx=11)
        self.sum_label.pack(side=tk.RIGHT)

    def _row4_pack(self):
        self.cashless_label.pack(side=tk.LEFT)
        self.cashless.pack(side=tk.LEFT, padx=16)
        self.pay_conditions_label.pack(side=tk.LEFT)
        self.pay_conditions_box.pack(side=tk.LEFT, padx=5)
        self.initiator_label.pack(side=tk.LEFT, padx=5)
        self.initiator_label_name.pack(side=tk.LEFT, padx=5)
        self.approval_label.pack(side=tk.LEFT)
        self.approval_box.pack(side=tk.LEFT, padx=5)

    def _top_pack(self):
        self.main_label.pack(side=tk.TOP, expand=False, anchor=tk.NW)
        self.limit_label.pack(side=tk.LEFT, expand=False, anchor=tk.W)
        self.limit_month.pack(side=tk.LEFT, expand=False, anchor=tk.W)
        self.limit_sum.pack(side=tk.LEFT, expand=False, anchor=tk.W)

    def _validate_request_creation(self, messagetitle, sumtotal):
        """ Check if all fields are filled properly. """
        if not self.mvz_current.get():
            messagebox.showerror(
                    messagetitle, 'Не указано МВЗ'
            )
            return False
        if not self.office_box.get():
            messagebox.showerror(
                    messagetitle, 'Не выбран офис'
            )
            return False
        if not self.category_box.get():
            messagebox.showerror(
                    messagetitle, 'Не выбрана категория'
            )
            return False
        if not self.pay_conditions_box.get():
            messagebox.showerror(
                    messagetitle, 'Не выбраны условия оплаты'
            )
            return False
        if not sumtotal:
            messagebox.showerror(
                    messagetitle, 'Не указана сумма'
            )
            return False
        date_validation = self._validate_plan_date()
        if date_validation == 'incorrect_date':
            messagebox.showerror(
                    messagetitle,
                    'Плановая дата не может быть прошедшей'
            )
            return False
        elif date_validation == 'ask_confirmation':
            confirmed = messagebox.askyesno(title='Подтвердите действие',
                            message='Создать заявку на сегодняшную дату?')
            if not confirmed:
                return False
        elif date_validation != 'correct_date':
            messagebox.showerror(title='Ошибка',
                            message=('Возникло непредвиденное исключение\n{}'
                                     .format(date_validation))
                                )
            return False
        return True

    def _validate_plan_date(self):
        """ Validate date correctness according to rules. """
        try:
            plan_date = self.plan_date_entry.get()
            try:
                plan_date = datetime.strptime(plan_date, '%d.%m.%Y').date()
            except ValueError:
                plan_date = datetime.strptime(plan_date, '%d.%m.%y').date()
            today = date.today()
            return ('correct_date' if plan_date > today else
                    'ask_confirmation' if plan_date == today else
                    'incorrect_date'
                    )
        except Exception as e:
            return e


class PreviewForm(PaymentFrame):
    def __init__(self, parent, controller, connection, user_info, mvz,
                 allowed_initiators, categories, pay_conditions, status_list, **kwargs):
        super().__init__(parent, controller, connection, user_info,
                         mvz, categories, pay_conditions)
        self.office = tuple(sorted(set(x for lst in map(lambda v: v[1],
                                                        self.mvz.values()) for x in lst),
                                   key=lambda s: '' if s == 'Все' else s))
        self.initiatorsID, self.initiators = zip(*allowed_initiators)
        self.statusID, self.status_list = zip(*[(None, 'Все'),] + status_list)
        # EXTENDED_MODE activates extended selectmode for treeview, realized
        # using checkboxes, and allows to approve multiple requests
        self.EXTENDED_MODE = (True if user_info.isSuperUser
                              and user_info.AccessType == 2 else False)
        # List of functions to get payments
        # determines what payments will be shown when refreshing
        self.payments_load_list = [self._get_all_payments,
                                   self._get_payments_for_approval]
        self.get_payments = self._get_all_payments
        # Parameters for sorting
        self.rows = None  # store all rows for sorting and redrawing
        self.sort_reversed_index = None  # reverse sorting for the last sorted column
        self.month = list(month_name)
        self.month_default = self.month[datetime.now().month]

        # Top Frame with description and user name
        top = tk.Frame(self, name='top_cf', padx=5)

        main_label = tk.Label(top, text='Просмотр заявок',
                              padx=10, font=('Arial', 8, 'bold'))
        main_label.pack(side=tk.LEFT, expand=False, anchor=tk.NW)

        self._add_copyright(top)
        self._add_user_label(top)

        top.pack(side=tk.TOP, fill=tk.X, expand=False)

        # Filters
        filterframe = ttk.LabelFrame(self, text=' Фильтры ', name='filterframe')

        # First Filter Frame with (MVZ, office)
        row1_cf = tk.Frame(filterframe, name='row1_cf', padx=15)

        self.initiator_label = tk.Label(row1_cf, text='Инициатор', padx=10)
        self.initiator_box = ttk.Combobox(row1_cf, width=35, state='readonly')
        self.initiator_box['values'] = self.initiators

        self.mvz_label = tk.Label(row1_cf, text='МВЗ', padx=10)
        self.mvz_box = ttk.Combobox(row1_cf, width=45, state='readonly')
        self.mvz_box['values'] = list(self.mvz)

        self.office_label = tk.Label(row1_cf, text='Офис', padx=20)
        self.office_box = ttk.Combobox(row1_cf, width=20, state='readonly')
        self.office_box['values'] = self.office

        self.status_label = tk.Label(row1_cf, text='Статус', padx=20)
        self.status_box = ttk.Combobox(row1_cf, width=10, state='readonly')
        self.status_box['values'] = self.status_list

        # Pack row1_cf
        self._row1_pack()
        row1_cf.pack(side=tk.TOP, fill=tk.X)

        # Second Fill Frame with (Plan date, Sum, Tax)
        row2_cf = tk.Frame(filterframe, name='row2_cf', padx=15)

        self.date_label = tk.Label(row2_cf, text='Дата', padx=10)
        self.date_type = ttk.Combobox(row2_cf, width=12, state='readonly')
        self.date_type['values'] = ['плановая', 'создания']
        self.date_label_m = tk.Label(row2_cf, text=':  месяц', padx=10)
        self.date_entry_m = MultiselectMenu(row2_cf, self.month_default,
                                                 self.month, width=15)
        self.date_label_y = tk.Label(row2_cf, text='год', padx=20)
        self.year = tk.IntVar()
        self.date_entry_y = tk.Spinbox(row2_cf, width=7, from_=2019, to=2059,
                                            font=('Arial', 9), textvariable=self.year)
        self.sum_label_from = tk.Label(row2_cf, text='Сумма без НДС: от')
        self.sumtotal_from = StringSumVar()
        vcmd = (self.register(self._validate_sum))
        self.sum_entry_from = tk.Entry(row2_cf, width=12, textvariable=self.sumtotal_from,
                        validate='all', validatecommand=(vcmd, '%P')
                        )
        self.sum_entry_from.bind("<FocusIn>", self._on_focus_in_format_sum)
        self.sum_entry_from.bind("<FocusOut>", self._on_focus_out_format_sum)
        self.sum_label_to = tk.Label(row2_cf, text='до')
        self.sumtotal_to = StringSumVar()
        self.sum_entry_to = tk.Entry(row2_cf, width=12, textvariable=self.sumtotal_to,
                        validate='all', validatecommand=(vcmd, '%P'))
        self.sum_entry_to.bind("<FocusIn>", self._on_focus_in_format_sum)
        self.sum_entry_to.bind("<FocusOut>", self._on_focus_out_format_sum)
        self.nds_label = tk.Label(row2_cf, text='НДС', padx=20)
        self.nds = tk.IntVar()
        self.ndsall = ttk.Radiobutton(row2_cf, text="Любой", variable=self.nds, value=-1)
        self.nds20 = ttk.Radiobutton(row2_cf, text="20 %", variable=self.nds, value=20)
        self.nds7 = ttk.Radiobutton(row2_cf, text="7 %", variable=self.nds, value=7)
        self.nds0 = ttk.Radiobutton(row2_cf, text="0 %", variable=self.nds, value=0)

        # Pack row2_cf
        self._row2_pack()
        row2_cf.pack(side=tk.TOP, fill=tk.X)
        filterframe.pack(side=tk.TOP, fill=tk.X, expand=False, padx=10, pady=5)

        # Third Fill Frame (checkbox + button to apply filter)
        row3_cf = tk.Frame(filterframe, name='row3_cf', padx=15)

        self.search_by_num = tk.StringVar()
        self.search_by_num_label = tk.Label(row3_cf, text='Поиск по номеру заявки')
        self.search_by_num_entry = tk.Entry(row3_cf, width=20,
                                            textvariable=self.search_by_num)

        self.bt3_0 = ttk.Button(row3_cf, text="Показать все заявки на утверждение",
                           width=40, command=self._show_payments_for_approval)
        self.bt3_1 = ttk.Button(row3_cf, text="Применить фильтр", width=20,
                         command=self._use_filter_and_refresh)
        self.bt3_2 = ttk.Button(row3_cf, text="Очистить фильтр", width=20,
                         command=self._clear_filters)

        # Pack row3_cf
        self._row3_pack()
        row3_cf.pack(side=tk.TOP, fill=tk.X, pady=10)

        if self.EXTENDED_MODE:
            # Fourth Frame (toggle select all rows, button to approve selected)
            row4_cf = tk.Frame(filterframe, name='row4_cf', padx=15)
            self.all_rows_checked = tk.IntVar()
            self.check_all_rows = tk.Checkbutton(row4_cf, text="Выбрать все",
                                      variable=self.all_rows_checked,
                                      command=self._toggle_all_rows)
            self.check_all_rows.pack(side=tk.LEFT)
            bt2b = ttk.Button(row4_cf, text="Утвердить выбранные", width=25,
                             command=self._approve_multiple)
            bt2b.pack(side=tk.LEFT, padx=10, pady=10)
            row4_cf.pack(side=tk.TOP, fill=tk.X)

        # Set all filters to default
        self._clear_filters()

        # Text Frame
        preview_cf = ttk.LabelFrame(self, text=' Заявки ', name='preview_cf')

        # column name and width
        #self.headings=('a', 'bb', 'cccc')  # for debug
        self.headings = {'№ п/п': 30, 'ID': 0, 'InitiatorID': 0, '№ заявки': 100,
            'Кем создано': 130, 'Дата создания': 80, 'Дата/время создания': 120,
            'CSP':30, 'МВЗ SAP': 70, 'МВЗ': 150, 'Офис': 80, 'Категория': 80,
            'Условия оплаты':80, 'Контрагент': 60, 'Плановая дата': 90,
            'Сумма без НДС': 85, 'Сумма с НДС': 85, 'Вид платежа':0, 'Статус': 45,
            'Статус заявки': 120, 'Описание': 120, 'ID Утверждающего': 0,
            'Утверждающий': 120}

        if self.EXTENDED_MODE:
            self.table = CheckboxTreeview(preview_cf,
                                          selectmode='browse',
                                          style='HeaderStyle.Treeview'
                                          )
        else:
            self.table = ttk.Treeview(preview_cf, show='headings',
                                      selectmode='browse',
                                      style='HeaderStyle.Treeview'
                                      )

        self._init_table(preview_cf)
        self.table.pack(expand=tk.YES, fill=tk.BOTH)

        # asserts for headings used through script as indices
        head = self.table["columns"]
        msg = 'Heading order must be reviewed. Wrong heading: '
        assert head[1] == 'ID', '{}ID'.format(msg)
        assert head[-2] == 'ID Утверждающего', '{}ID Утверждающего'.format(msg)
        assert head[-5] == 'Статус', '{}Статус'.format(msg)
        assert head[-7] == 'Сумма с НДС', '{}Сумма с НДС'.format(msg)

        # Bottom Frame with buttons
        bottom_cf = tk.Frame(self, name='bottom_cf')
        # Show create buttons only for users with rights
        # to create (1) or approve + create (2)
        if self.user_info.AccessType in (1, 2):
            bt1 = ttk.Button(bottom_cf, text="Создать", width=10,
                             command=lambda: controller._show_frame('CreateForm'))
            bt1.pack(side=tk.LEFT, padx=10, pady=10)

            bt2 = ttk.Button(bottom_cf, text="Создать из заявки", width=20,
                             command=self._create_from_current)
            bt2.pack(side=tk.LEFT, padx=10, pady=10)

        if self.user_info.UserID in (42, 76, 20):
            bt3 = ttk.Button(bottom_cf, text="Изменить заявку", width=20,
                             command=self._alter_request)
            bt3.pack(side=tk.LEFT, padx=10, pady=10)

        bt6 = ttk.Button(bottom_cf, text="Выход", width=10,
                         command=controller._quit)
        bt6.pack(side=tk.RIGHT, padx=10, pady=10)

        bt5 = ttk.Button(bottom_cf, text="Подробно", width=10,
                         command=self._show_detail)
        bt5.pack(side=tk.RIGHT, padx=10, pady=10)

        bt4 = ttk.Button(bottom_cf, text="Экспорт в Excel", width=15,
                         command=self._export_to_excel)
        bt4.pack(side=tk.RIGHT, padx=10, pady=10)

        if self.userID in (6, 24, 42, 76, 20):
            bt4a = ttk.Button(bottom_cf, text="Изменить лимиты", width=20,
                command=self._alter_limits)
            bt4a.pack(side=tk.RIGHT, padx=10, pady=10)
            bt4b = ttk.Button(bottom_cf, text="Открыть отчёт", width=20,
                command=self._open_report)
            bt4b.pack(side=tk.RIGHT, padx=10, pady=10)

        # Pack frames
        bottom_cf.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        preview_cf.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

    def _add_copyright(self, parent):
        """ Adds user name in the top right corner. """
        copyright_label = tk.Label(parent, text="О программе",
                                   font=('Arial', 8, 'underline'),
                                   cursor="hand2")
        copyright_label.bind("<Button-1>", self._show_about)
        copyright_label.pack(side=tk.RIGHT, anchor=tk.N)

    def _alter_limits(self):
        """ Create and raise new frame with limits. """
        self._raise_Toplevel(frame=AlterLimits, title='Изменение лимитов',
                             width=400, height=300, static_geometry=False,
                             options=(self.conn,))

    def _alter_request(self):
        """ Create and raise new frame to alter chosen payment. """
        curRow = self.table.focus()
        if curRow:
            request = dict(zip(self.table["columns"],
                               self.table.item(curRow).get('values')))
            request_info = self.conn.get_info_to_alter_payment(request['ID'])
            if not request_info:
                messagebox.showinfo(
                        'Изменение заявки',
                        'Изменять можно только утверждённые заявки.'
                )
                return
            self._raise_Toplevel(frame=AlterRequest, title='Изменение заявки',
                                 width=200, height=180, static_geometry=False,
                                 options=(self, self.conn, self.userID,
                                          request_info[0]))

    def _approve_multiple(self):
        """ Allows to approve multiple requests chosen in PreviewForm.
        """
        curItems = (item for item in self.table.get_children()
                    if self.table.tag_has('checked', item))
        if not curItems:
            return
        # store paymentID and SumNoTax
        to_approve = {}
        for curRow in curItems:
            request = self.table.item(curRow).get('values')
            # extract all approvable requests for current user
            if self._is_valid_approval(request[-2]):
                to_approve[request[1]] = float(request[-7].replace(' ', '')
                                                          .replace(',', '.'))
        if not to_approve:
            return
        appr_sum = '{:,.0f}'.format(sum(to_approve.values())).replace(',', ' ')
        confirmed = messagebox.askyesno(title='Подтвердите действие',
            message='Выбрано заявок: {} на сумму {} грн.\nУтвердить заявки?'
                    .format(len(to_approve), appr_sum)
                                       )
        if not confirmed:
            return
        for paymentID in to_approve:
            self.conn.update_confirmed(self.userID, paymentID, is_approved=True)
        self._refresh()

    def _center_popup_window(self, newlevel, w, h, static_geometry=True):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        start_x = int((screen_width/2) - (w/2))
        start_y = int((screen_height/2) - (h * 0.7))

        if static_geometry == True:
            newlevel.geometry('{}x{}+{}+{}'.format(w, h, start_x, start_y))
        else:
            newlevel.geometry('+{}+{}'.format(start_x, start_y))

    def _change_preview_state(self, new_state):
        """ Change payments state that determines which payments will be shown.
        """
        if new_state == 'Show payments according to filters':
            self.get_payments = self._get_all_payments
        elif new_state == 'Show payments for approval':
            self.get_payments = self._get_payments_for_approval

    def _check_rights_to_fill_CreateForm(self, to_fill):
        try:
            allowed_offices = self.get_offices(to_fill['МВЗ'])
            if to_fill['Офис'] not in allowed_offices:
                raise NoRightsToFillCreateFormError(to_fill['Офис'])
        except KeyError:
            raise NoRightsToFillCreateFormError(to_fill['МВЗ'])

    def _clear_filters(self):
        self.initiator_box.set('Все')
        self.mvz_box.set('Все')
        self.office_box.set('Все')
        self.status_box.set('Все')
        self.date_type.set('плановая')
        self.date_entry_m.set_default_option()
        self.year.set(datetime.now().year)
        self.sumtotal_from.set('0,00')
        self.sumtotal_to.set('')
        self.nds.set(-1)
        self.search_by_num.set('')

    def _create_from_current(self):
        """ Raises CreateForm with partially filled labels/entries. """
        curRow = self.table.focus()
        if curRow:
            # extract info to be putted in CreateForm
            to_fill = dict(zip(self.table["columns"],
                               self.table.item(curRow).get('values')))
            try:
                self._check_rights_to_fill_CreateForm(to_fill)
            except NoRightsToFillCreateFormError as e:
                messagebox.showerror(self.controller.title(),
                                     e.message + '\n' + e.expression)
                return
            self.controller._fill_CreateForm(**to_fill)
            self.controller._show_frame('CreateForm')

    def _export_to_excel(self):
        if not self.rows:
            return
        headings = {k:v for k, v in self.headings.items() if k != '№ п/п'}
        isExported = export_to_excel(headings, self.rows)
        if isExported:
            messagebox.showinfo(
                'Экспорт в Excel',
                'Данные экспортированы на рабочий стол'
            )
        else:
            messagebox.showerror(
                'Экспорт в Excel',
                'При экспорте произошла непредвиденная ошибка'
            )

    def _get_all_payments(self):
        """ Extract information from filters and get payments list. """
        filters = {'initiator': self.initiatorsID[self.initiator_box.current()],
                   'mvz': self.get_mvzSAP(self.mvz_box.get()),
                   'office': (self.office_box.current() and
                              self.office[self.office_box.current()]),
                   'date_type': self.date_type.current(),
                   'date_m': self.date_entry_m.get_selected(),
                   'date_y': self.year.get() if self.date_entry_y.get() else 0.,
                   'sumtotal_from': float(self.sumtotal_from.get_float_form()
                                          if self.sum_entry_from.get() else 0),
                   'sumtotal_to': float(self.sumtotal_to.get_float_form()
                                        if self.sum_entry_to.get() else 0),
                   'nds':  self.nds.get(),
                   'statusID': (self.status_box.current() and
                              self.statusID[self.status_box.current()]),
                   'payment_num': self.search_by_num.get().strip()
                   }
        if not filters['date_m']:
            raise MonthFilterError(filters['date_m'])
        self.rows = self.conn.get_paymentslist(user_info=self.user_info,
                                               **filters)

    def _get_payments_for_approval(self):
        """ Get info about payments nedd to be approved. """
        self.rows = self.conn.get_paymentslist(user_info=self.user_info,
                                               for_approval=True)

    def _is_valid_approval(self, approvalID):
        """ Check if current user is approval person.
            Treat users with ID 24 and 9 as the same person.
        """
        userID = (9 if self.userID == 24 else self.userID)
        approvalID = (9 if approvalID == 24 else approvalID)
        return (userID == approvalID and (self.user_info.AccessType == 2
                                          or self.user_info.isSuperUser == 1))

    def _init_table(self, parent):
        """ Creates treeview. """
        if isinstance(self.headings, dict):
            self.table["columns"] = tuple(self.headings.keys())
            self.table["displaycolumns"] = tuple(k for k in self.headings.keys()
                if k not in ('ID', 'НДС', 'Описание', 'Дата/время создания',
                             'МВЗ', 'Категория', 'Условия оплаты', 'Вид платежа',
                             'Статус заявки','InitiatorID', 'ID Утверждающего'))
            for head, width in self.headings.items():
                self.table.heading(head, text=head, anchor=tk.CENTER)
                self.table.column(head, width=width, anchor=tk.CENTER)

        else:
            self.table["columns"] = self.headings
            self.table["displaycolumns"] = self.headings
            for head in self.headings:
                self.table.heading(head, text=head, anchor=tk.CENTER)
                self.table.column(head, width=50*len(head), anchor=tk.CENTER)

        # for debug
        #self._show_rows(rows=((123, 456, 789), ('abc', 'def', 'ghk')))

        for tag, bg in zip(self.status_list[1:5],
                           ('#ffffc8', 'lightgray', 'lightgreen', '#f66e6e')):
            self.table.tag_configure(tag, background=bg)
        #self.table.tag_configure('oddrow', background='lightgray')

        self.table.bind('<Double-1>', self._show_detail)
        self.table.bind('<Button-1>', self._sort, True)

        scrolltable = tk.Scrollbar(parent, command=self.table.yview)
        self.table.configure(yscrollcommand=scrolltable.set)
        scrolltable.pack(side=tk.RIGHT, fill=tk.Y)

    def _open_report(self):
        """ Open independent report. """
        os.startfile(os.path.join(REPORT_PATH, 'Заявки_на_оплату.xlsb'))

    def _raise_Toplevel(self, frame, title, width, height,
                        static_geometry=True, options=()):
        """ Create and raise new frame with limits.
        Input:
        frame - class, Frame class to be drawn in Toplevel;
        title - str, window title;
        width - int, width parameter to center window;
        height - int, height parameter to center window;
        static_geometry - bool, if True - width and height will determine size
            of window, otherwise size will be determined dynamically;
        options - tuple, arguments that will be sent to frame.
        """
        newlevel = tk.Toplevel(self.parent)
        #newlevel.transient(self)  # disable minimize/maximize buttons
        newlevel.title(title)
        newlevel.bind('<Escape>', lambda e, w=newlevel: w.destroy())
        frame(newlevel, *options)
        newlevel.resizable(width=False, height=False)
        self._center_popup_window(newlevel, width, height, static_geometry)
        newlevel.focus()
        newlevel.grab_set()

    def _refresh(self):
        """ Refresh information about payments. """
        try:
            self.get_payments()
        except MonthFilterError as e:
            messagebox.showerror(self.controller.title(), e.message)
            return
        self._show_rows(self.rows)
        # Deselect "check_all_rows" checkbutton
        if self.EXTENDED_MODE:
            self.all_rows_checked.set(0)

    def _resize_columns(self):
        """ Resize columns in treeview. """
        self.table.column('#0', width=36)
        for head, width in self.headings.items():
            self.table.column(head, width=width)

    def _row1_pack(self):
        self.initiator_label.pack(side=tk.LEFT)
        self.initiator_box.pack(side=tk.LEFT, padx=5, pady=5)
        self.mvz_label.pack(side=tk.LEFT)
        self.mvz_box.pack(side=tk.LEFT, padx=5, pady=5)
        self.status_box.pack(side=tk.RIGHT, padx=5, pady=5)
        self.status_label.pack(side=tk.RIGHT)
        self.office_box.pack(side=tk.RIGHT, padx=5, pady=5)
        self.office_label.pack(side=tk.RIGHT)

    def _row2_pack(self):
        self.date_label.pack(side=tk.LEFT)
        self.date_type.pack(side=tk.LEFT)
        self.date_label_m.pack(side=tk.LEFT)
        self.date_entry_m.pack(side=tk.LEFT)
        self.date_label_y.pack(side=tk.LEFT)
        self.date_entry_y.pack(side=tk.LEFT, anchor=tk.SW, pady=10)
        self.nds0.pack(side=tk.RIGHT, padx=7)
        self.nds7.pack(side=tk.RIGHT, padx=7)
        self.nds20.pack(side=tk.RIGHT, padx=7)
        self.ndsall.pack(side=tk.RIGHT, padx=7)
        self.nds_label.pack(side=tk.RIGHT, padx=5)
        self.sum_entry_to.pack(side=tk.RIGHT)
        self.sum_label_to.pack(side=tk.RIGHT, padx=2)
        self.sum_entry_from.pack(side=tk.RIGHT)
        self.sum_label_from.pack(side=tk.RIGHT, padx=2)

    def _row3_pack(self):
        self.bt3_0.pack(side=tk.LEFT, padx=10)
        self.bt3_2.pack(side=tk.RIGHT, padx=10)
        self.bt3_1.pack(side=tk.RIGHT, padx=10)
        self.search_by_num_entry.pack(side=tk.RIGHT, padx=10)
        self.search_by_num_label.pack(side=tk.RIGHT, padx=10)

    def _show_about(self, event=None):
        """ Raise frame with info about app. """
        self._raise_Toplevel(frame=AboutFrame,
                             title='Заявки на оплату v. ' + __version__,
                             width=400, height=150)

    def _show_detail(self, event=None):
        """ Show details when double-clicked on row. """
        show_detail = (not event or (self.table.identify_row(event.y) and
                                     int(self.table.identify_column(event.x)[1:]) > 0
                                     )
                      )
        if show_detail:
            curRow = self.table.focus()
            if curRow:
                newlevel = tk.Toplevel(self.parent)
                newlevel.withdraw()
                #newlevel.transient(self.parent)  # disable minimize/maximize buttons
                newlevel.title('Заявка детально')
                newlevel.iconbitmap('resources/preview.ico')
                newlevel.bind('<Escape>', lambda e, w=newlevel: w.destroy())
                approvalID = self.table.item(curRow).get('values')[-2]
                is_valid_approval = self._is_valid_approval(approvalID)
                if is_valid_approval:
                    ApproveConfirmation(newlevel, self, self.conn, self.userID,
                                        self.headings,
                                        self.table.item(curRow).get('values'),
                                        self.table.item(curRow).get('tags'))
                else:
                    DetailedPreview(newlevel, self, self.conn, self.userID,
                                    self.headings,
                                    self.table.item(curRow).get('values'),
                                    self.table.item(curRow).get('tags'))
                newlevel.resizable(width=False, height=False)
                # width is set implicitly in DetailedPreview._newRow
                # based on columnWidths values
                self._center_popup_window(newlevel, 500, 400,
                                          static_geometry=False)
                newlevel.deiconify()
                newlevel.focus()
                newlevel.grab_set()
        else:
            # if double click on header - redirect to sorting rows
            self._sort(event)

    def _sort(self, event):
        if self.table.identify_region(event.x, event.y) == 'heading' and self.rows:
            # determine index of displayed column
            disp_col = int(self.table.identify_column(event.x)[1:]) - 1
            if disp_col < 1:  # ignore sort by '№ п/п' and checkboxes
                return
            # determine index of this column in self.rows
            # substract 1 because of added '№ п/п' which don't exist in data
            sort_col = self.table["columns"].index(self.table["displaycolumns"][disp_col]) - 1
            self.rows.sort(key=lambda x: x[sort_col],
                           reverse=self.sort_reversed_index == sort_col)
            # store index of last sorted column if sort wasn't reversed
            self.sort_reversed_index = None if self.sort_reversed_index==sort_col else sort_col
            self._show_rows(self.rows)

    def _show_payments_for_approval(self):
        """ Change state to show payments for approval. """
        self._change_preview_state('Show payments for approval')
        self._refresh()

    def _show_rows(self, rows):
        """ Refresh table with new rows. """
        self.table.delete(*self.table.get_children())
        if not rows:
            return
        for i, row in enumerate(rows):
            # tag = (Status, 'unchecked')
            self.table.insert('', tk.END,
                values=(i+1,) + tuple(map(lambda val: self._format_float(val)
                if isinstance(val, Decimal) else val, row)),
                tags=(row[-5], 'unchecked'))

    def _toggle_all_rows(self, event=None):
        if self.all_rows_checked.get():
            for item in self.table.get_children():
                self.table.check_item(item)
        else:
            for item in self.table.get_children():
                self.table.uncheck_item(item)

    def _use_filter_and_refresh(self):
        """ Change state to filter usage. """
        self._change_preview_state('Show payments according to filters')
        self._refresh()


class DetailedPreview(tk.Frame):
    """ Class that creates Frame with information about chosen request. """
    def __init__(self, parent, parentform, conn, userID, head, info, tags):
        super().__init__(parent)
        self.parent = parent
        self.parentform = parentform
        self.conn = conn
        self.approveclass_bool = isinstance(self, ApproveConfirmation)
        self.paymentID, self.initiatorID = info[1:3]
        self.userID = userID
        self.rowtags = tags

        # Top Frame with description and user name
        self.top = tk.Frame(self, name='top_cf', padx=5, pady=5)

        # Create a frame on the canvas to contain the buttons.
        self.table_frame = tk.Frame(self.top)

        # Add info to table_frame
        fonts = (('Arial', 9, 'bold'), ('Arial', 10))
        for row in zip(range(len(head)), zip(head, info)):
            if row[1][0] not in ('№ п/п', 'ID', 'InitiatorID', 'Дата создания',
                                 'ID Утверждающего', 'Утверждающий', 'Статус'):
                self._newRow(self.table_frame, fonts, *row)

        self.appr_label = tk.Label(self.top, text='Утверждающие',
                                   padx=10, pady=5, font=('Arial', 10, 'bold'))

        # Top Frame with description and user name
        self.appr_cf = tk.Frame(self, name='appr_cf', padx=5)

        # Add approvals to appr_cf
        fonts = (('Arial', 10), ('Arial', 10))
        approvals = self.conn.get_approvals(self.paymentID)
        for rowNumber, row in enumerate(approvals):
            self._newRow(self.appr_cf, fonts, rowNumber+1, row)

        self._add_buttons()
        self._pack_frames()

    def _add_buttons(self):
        # Bottom Frame with buttons
        self.bottom = tk.Frame(self, name='bottom')

        if self.approveclass_bool:
            bt1 = ttk.Button(self.bottom, text="Утвердить", width=10,
                             command=lambda: self._close(True),
                             style='ButtonGreen.TButton')
            bt1.pack(side=tk.LEFT, padx=15, pady=5)

            bt2 = ttk.Button(self.bottom, text="Отклонить", width=10,
                             command=lambda: self._close(False),
                             style='ButtonRed.TButton')
            bt2.pack(side=tk.LEFT, padx=15, pady=5)

        bt4 = ttk.Button(self.bottom, text="Закрыть", width=10,
                         command=self.parent.destroy)
        bt4.pack(side=tk.RIGHT, padx=15, pady=5)

        if self.userID == self.initiatorID and 'Отозв.' not in self.rowtags:
            bt3 = ttk.Button(self.bottom, text="Отозвать", width=10,
                             command=self._discard)
            bt3.pack(side=tk.RIGHT, padx=15, pady=5)

    def _discard(self):
        mboxname = 'Отзыв заявки'
        confirmed = messagebox.askyesno(title=mboxname,
                  message='Вы уверены, что хотите отозвать заявку?')
        if confirmed:
            self.conn.update_discarded(self.paymentID)
            messagebox.showinfo(mboxname, 'Заявка отозвана')
            self.parentform._refresh()
            self.parent.destroy()

    def _newRow(self, frame, fonts, rowNumber, info):
        """ Adds a new line to the table. """

        numberOfLines = []       # List to store number of lines needed
        columnWidths = [20, 50]  # Width of the different columns in the table

        # Find the length and the number of lines of each element and column
        for index, item in enumerate(info):
            # minimum 1 line + number of new lines + lines that too long
            numberOfLines.append(1 + str(item).count('\n') +
                sum(floor(len(s)/columnWidths[index]) for s in str(item).split('\n'))
            )

        # Find the maximum number of lines needed
        lineNumber = max(numberOfLines)

        # Define labels (columns) for row
        def form_column(rowNumber, lineNumber, col_num, cell, fonts):
            col = tk.Text(frame, bg='white', padx=3)
            col.insert(1.0, cell)
            col.grid(row=rowNumber, column=col_num+1, sticky='news')
            col.configure(width=columnWidths[col_num], height=min(9, lineNumber),
                          font=fonts[col_num], state="disabled")
            if lineNumber > 9 and col_num == 1:
                scrollbar = tk.Scrollbar(frame, command=col.yview)
                scrollbar.grid(row=rowNumber, column=col_num+2, sticky='nsew')
                col['yscrollcommand'] = scrollbar.set

        for col_num, cell in enumerate(info):
            form_column(rowNumber, lineNumber, col_num, cell, fonts)

    def _pack_frames(self):
        self.top.pack(side=tk.TOP, fill=tk.X, expand=False)
        self.bottom.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.appr_cf.pack(side=tk.TOP, fill=tk.X)
        self.table_frame.pack()
        self.appr_label.pack(side=tk.LEFT, expand=False)
        self.pack()


class ApproveConfirmation(DetailedPreview):
    """ Class with information about reuqest that contains buttons
        to approval/decline it. """
    def __init__(self, parent, parentform, conn, userID, head, info, tags):
        super().__init__(parent, parentform, conn, userID, head, info, tags)

    def _close(self, is_approved):
        confirmed = messagebox.askyesno(title='Подтвердите действие',
            message='{} заявку?'.format('Утвердить' if is_approved else 'Отклонить'))
        if confirmed:
            self.conn.update_confirmed(self.userID, self.paymentID, is_approved)
            self.parentform._refresh()
            self.parent.destroy()


class AboutFrame(tk.Frame):
    """ Creates a frame with copyright and info about app. """
    def __init__(self, parent):
        super().__init__(parent)

        self.top = ttk.LabelFrame(self, name='top_af')

        logo = tk.PhotoImage(file='resources/payment.png')
        self.logo_label = tk.Label(self.top, image=logo)
        self.logo_label.image = logo  # keep a reference to avoid gc!

        self.copyright_text = tk.Text(self.top, bg='#f1f1f1',
                                      font=('Arial', 8), relief=tk.FLAT)
        hyperlink = HyperlinkManager(self.copyright_text)

        def link_instruction():
            path = 'resources\\README.pdf'
            os.startfile(path)

        self.copyright_text.insert(tk.INSERT,
                                  'Заявки на оплату v. ' + __version__ +'\n\n')
        self.copyright_text.insert(tk.INSERT, "Инструкция по использованию",
                                   hyperlink.add(link_instruction))
        self.copyright_text.insert(tk.INSERT, "\n\n")

        def link_license():
            path = 'resources\\LICENSE.txt'
            os.startfile(path)

        self.copyright_text.insert(tk.INSERT,
                            'Copyright © 2019 Офис контроллинга логистики\n')
        self.copyright_text.insert(tk.INSERT, 'MIT License',
                                   hyperlink.add(link_license))

        self.bt = ttk.Button(self, text="Закрыть", width=10,
                        command=parent.destroy)
        self.pack_all()

    def pack_all(self):
        self.bt.pack(side=tk.BOTTOM, pady=5)
        self.top.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10)
        self.logo_label.pack(side=tk.LEFT, padx=10)
        self.copyright_text.pack(side=tk.LEFT, padx=10)
        self.pack(fill=tk.BOTH, expand=True)


class AlterLimits(tk.Frame):
    """ Creates a frame to manage user limits. """
    def __init__(self, parent, conn):
        super().__init__(parent)
        self.parent = parent
        self.conn = conn
        self.limits = self.conn.get_limits_info()
        # {Column name: width}
        self.headings = {'UserID': 6, 'Инициатор': 41,
                         'Лимит': 9, 'Обнулять': 8}

        # Top Frame with table
        self.top = tk.Frame(self, name='top_al')
        for j, (header, width) in enumerate(self.headings.items()):
            lb = tk.Label(self.top, text=header, font=('Arial', 8, 'bold'),
                          relief='sunken', borderwidth=1, width=width)
            lb.grid(row=0, column=j, ipadx=5, sticky='nswe')

        # Top Canvas Frame to connect Frame and Scrollbar
        self.canvas = tk.Canvas(self, borderwidth=0, background="#ffffff",
                                width=500, height=200)

        self.table = LabelGrid(self.canvas,
                               content=self.limits,
                               grid_width = (8, 42, 12, 50)
                               )

        self.scrolltable = tk.Scrollbar(self, orient="vertical",
                                        command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrolltable.set)
        self.canvas.create_window((4,4), window=self.table, anchor="nw",
                                  tags="self.table")

        self.table.bind("<Configure>", self._onFrameConfigure)

        # Bottom Frame with buttons
        self.bottom = tk.Frame(self, name='bottom_al')

        bt2 = ttk.Button(self.bottom, text="Закрыть", width=10,
                         command=self.parent.destroy)
        bt2.pack(side=tk.RIGHT, padx=15, pady=5)

        bt1 = ttk.Button(self.bottom, text="Сохранить", width=10,
                         command=self._update)
        bt1.pack(side=tk.RIGHT, padx=15, pady=5)
        self._pack_frames()

    def _onFrameConfigure(self, event):
        """ Reset the scroll region to encompass the inner frame. """
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _pack_frames(self):
        self.top.pack(side=tk.TOP, fill=tk.X, expand=False, padx=4, pady=2)
        self.bottom.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.scrolltable.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.pack()

    def _update(self):
        """ Update information about limits on server.
        """
        messagetitle = self.parent.title()
        try:
            limits = self.table.get_values()
        except ValueError:
            messagebox.showerror(
                    messagetitle, 'Введена некорректная сумма'
            )
            return
        update_success = self.conn.update_limits(limits)
        if update_success:
            messagebox.showinfo(
                    messagetitle, 'Изменения внесены'
            )
            self.parent.destroy()
        else:
            messagebox.showerror(
                    messagetitle, 'Произошла непредвиденная ошибка'
            )


class AlterRequest(tk.Frame):
    """ Creates a frame to alter chosen request. """
    def __init__(self, parent, parentform, conn, userID, request_info):
        super().__init__(parent)
        self.parent = parent
        self.parentform = parentform
        self.conn = conn
        self.userID = userID
        self.paymentID, self.request_date_str, self.request_sum = request_info
        self.request_date = datetime.strptime(self.request_date_str,
                                              '%Y-%m-%d').date()

        # Top Frame with table
        self.top = tk.Frame(self, name='top_ar')
        self.top_label = tk.Label(self.top,
            text='Измените необходимое поле:', font=('Arial', 10, 'bold'))

        self.altdata_head_label = tk.Label(self.top,
            text='Изменение даты', font=('Arial', 9, 'bold'), padx=10)
        self.altdata_label = tk.Label(self.top,
            text=('Укажите новую дату +/- 7 дней в пределах месяца. '
                  'Первичная дата:'),
            font=('Arial', 9), padx=10)
        self.plan_date = tk.StringVar()
        self.plan_date_entry = DateEntry(self.top, width=12, state='readonly',
                                         textvariable=self.plan_date, font=('Arial', 9),
                                         selectmode='day', borderwidth=2,
                                         locale='ru_RU')
        self.plan_date_entry.set_date(self.request_date)

        self.altsum_head_label = tk.Label(self.top,
            text='Изменение суммы', font=('Arial', 9, 'bold'), padx=10)
        self.altsum_label = tk.Label(self.top,
            text='Укажите сумму без НДС. Первичная сумма:',
            font=('Arial', 9), padx=10)
        self.sumtotal = StringSumVar()
        self.altsum_entry = tk.Entry(self.top, width=12, textvariable=self.sumtotal)
        self.sumtotal.set(self.request_sum)

        # Bottom Frame with buttons
        self.bottom = tk.Frame(self, name='bottom_ar')

        bt2 = ttk.Button(self.bottom, text="Отменить", width=10,
                         style='ButtonRed.TButton',
                         command=self.parent.destroy)
        bt2.pack(side=tk.RIGHT, padx=15, pady=5)

        bt1 = ttk.Button(self.bottom, text="Сохранить", width=10,
                         style='ButtonGreen.TButton',
                         command=self._apply_changes)
        bt1.pack(side=tk.LEFT, padx=15, pady=5)
        self._pack_frames()

    def _apply_changes(self):
        try:
            self._validate_changes()
        except (IncorrectFloatError, PeriodExceededError,
                MonthChangedError, SumExceededError) as e:
            messagebox.showerror(self.parent.title(),
                                 e.message + '\n' + e.expression)
            return
        new_date = self.plan_date_entry.get_date()
        # convert to datetime for SQL
        new_date = datetime.combine(new_date, datetime.min.time())
        new_sum = Decimal(self.sumtotal.get_float_form()).quantize(Decimal('.01'))
        result = self.conn.alter_payment(self.userID, self.paymentID,
                                         new_date, new_sum)
        if not result:
            messagebox.showerror(self.parent.title(),
                                 'Произошла ошибка при выполнении запроса')
        else:
            messagebox.showinfo(self.parent.title(), 'Запрос выполнен')
            self.parentform._refresh()
            self.parent.destroy()

    def _pack_frames(self):
        self.top_label.pack(side=tk.TOP)
        self.altdata_head_label.pack(side=tk.TOP)
        self.altdata_label.pack(side=tk.TOP)
        self.plan_date_entry.pack(side=tk.TOP, pady=5)
        self.altsum_head_label.pack(side=tk.TOP)
        self.altsum_label.pack(side=tk.TOP)
        self.altsum_entry.pack(side=tk.TOP)
        self.top.pack(side=tk.TOP, fill=tk.X, expand=False, padx=4, pady=2)
        self.bottom.pack(side=tk.BOTTOM, fill=tk.X, expand=False)
        self.pack()

    def _validate_changes(self):
        new_date = self.plan_date_entry.get_date()
        # Input check
        try:
            new_sum = Decimal(self.sumtotal.get_float_form()).quantize(Decimal('.01'))
        except ValueError:
            raise IncorrectFloatError(self.sumtotal.get())
        date_diff = (self.request_date - new_date)
        # check if period is not exceeded
        if abs(date_diff.days) > 7:
            raise PeriodExceededError(str((date_diff.days,
                                          self.request_date,
                                          new_date))
                )
        # check if month is the same
        if self.request_date.month != new_date.month:
            raise MonthChangedError(str(self.request_date))
        # check if sum is not exceeded
        if new_sum > self.request_sum:
            raise SumExceededError(str(self.request_sum))


if __name__ == '__main__':
    from db_connect import DBConnect
    from collections import namedtuple

    UserInfo = namedtuple('UserInfo', ['UserID', 'ShortUserName',
                                       'AccessType', 'isSuperUser', 'GroupID',
                                       'PayConditionsID'])

    with DBConnect(server='s-kv-center-s59', db='AnalyticReports') as sql:
        try:
            app = PaymentApp(connection=sql,
                             user_info=UserInfo(24, 'TestName', 2, 1, 1, 2),
                             mvz=[('20511RC191', '20511RC191', 'Офис'),
                                  ('40900A2595', '40900A2595', 'Офис')],
                             categories=[('Cat1', 1), ('Cat2', 2)],
                             pay_conditions=[('Cat1', 1), ('Cat2 cat', 2)],
                             allowed_initiators=[(None, 'Все'), (1, 2), (3, 4)],
                             approvals_for_first_stage=[('a', 1), ('b', 2)],
                             status_list=[(1, 'На согл.'), (2, 'Отозв.')]
                             )
            app.mainloop()
        except Exception as e:
            print(e)
            raise
    input('Press Enter...')