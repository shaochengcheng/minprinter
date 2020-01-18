#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""The frontend GUI.
"""

import logging
import numbers
import os
from os.path import expanduser, join

from appJar import gui
from openpyxl import load_workbook

from .backend import print_invoice

logger = logging.getLogger('FrontendGUI')
logger.setLevel(logging.DEBUG)


def add_my_file_logging(
        logger,
        filename,
        filemode='w',
        log_level=logging.INFO,
        log_format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'):
    """Add self-configured logging file handler to a logger (probably root)
    """
    fh = logging.FileHandler(filename, mode=filemode)
    fh.setLevel(log_level)
    fm = logging.Formatter(log_format)
    fh.setFormatter(fm)
    logger.addHandler(fh)


class MPrinterGUI():
    """ A class to print invoices from china mobile or unicom.
    """
    def __init__(self, name='China Mobile and Unicom Invoices Printer'):
        self.app = gui(name)
        self.setting_keys = [
            'input_dir', 'output_dir', 'dpi', 'recursive', 'do_analysis'
        ]
        self.settings = dict(input_dir=None,
                             output_dir=None,
                             dpi=600,
                             recursive=True,
                             do_analysis=True)
        self.was_run = False
        self.last_settings = self.settings.copy()
        self.output_filenames = dict(stats='output-results.xlsx',
                                     pdf='output-jpg-a4.pdf')
        self.log_file = expanduser(join('~', 'minprinter-log.txt'))
        add_my_file_logging(logging.getLogger(), filename=self.log_file)

    def fill_input(self):
        """Fill the settings from input widgets.
        """
        self.settings['input_dir'] = self.app.getEntry('input_dir').strip()
        self.settings['output_dir'] = self.app.getEntry('output_dir').strip()
        self.settings['dpi'] = self.app.getSpinBox('DPI')
        self.settings['recursive'] = self.app.getCheckBox('Include Subdirs')
        self.settings['do_analysis'] = self.app.getCheckBox('Statistics')
        if self.settings['dpi']:
            self.settings['dpi'] = int(self.settings['dpi'])

    def run(self):
        """The main function for `Run` button.
        """
        self.fill_input()
        # precheck
        if not self.settings['input_dir']:
            self.app.errorBox('Error', 'Please select input PDFs folder!')
            return
        if not self.settings['output_dir']:
            self.app.errorBox('Error', 'Please select output folder!')
        # any changes
        if self.was_run is True:
            any_change_flag = False
            for k in self.setting_keys:
                if self.settings[k] != self.last_settings[k]:
                    any_change_flag = True
            if not any_change_flag:
                msg = 'No changes of settings since last run, you could find'\
                    + ' results in output directory {!r}\n.'\
                    + ' Would you like to run any way?'
                msg = msg.format(self.settings['output_dir'])
                if self.app.yesNoBox('Warning', msg) is False:
                    return
        # run background function
        try:
            print_invoice(input_dir=self.settings['input_dir'],
                          output_dir=self.settings['output_dir'],
                          output_filenames=self.output_filenames,
                          dpi=self.settings['dpi'],
                          recursive=self.settings['recursive'],
                          do_analysis=self.settings['do_analysis'])
        except Exception as err:
            self.app.errorBox('Error', str(err))
            raise
        # fresh frontend GUI
        self.app.openTab('TabbedFrame', 'Logs')
        with open(self.log_file) as fp:
            logs = fp.read()
        self.app.clearTextArea('Logs')
        self.app.setTextArea('Logs', logs)
        self.app.openTab('TabbedFrame', 'Results')
        #         additional_tip = """\n\r\n\r\n\r
        # Please check the detailed statistics results file {!r} in output\
        # directory {!r}""".format(self.output_filenames['stats'],
        #                          self.settings['output_dir'])
        #         self.app.message(title='Additional Info',
        #                          value=additional_tip,
        #                          width=750)
        self.app.openScrollPane('pane')
        try:
            excel_filename = join(self.settings['output_dir'],
                                  self.output_filenames['stats'])
            wb = load_workbook(excel_filename)
            ws = wb.worksheets[0]
            data = [
                list(
                    map(
                        lambda v: '{:.2f}'.format(v)
                        if isinstance(v, numbers.Number) else v, rv))
                for rv in ws.values
            ]
            self.app.table(title='results', value=data, width=800)
        except Exception as err:
            self.app.errorBox('Error', str(err))
            raise
        self.app.setTabbedFrameSelectedTab('TabbedFrame', 'Results')
        self.was_run = True
        self.last_settings = self.settings.copy()
        self.app.stopScrollPane()
        self.app.stopTabbedFrame()

    def on_input_dir_change(self):
        """Function when DirectoryEntry `input_dir` changes.
        """
        if not self.app.getEntry('output_dir'):
            self.app.setEntry('output_dir', self.app.getEntry('input_dir'))

    def draw_app(self):
        """ Draw the GUI of our app. Note functions are not bind correctly yet.
        """
        self.app.setBg('#FFFFFF')
        # app.setTransparency(0.8)
        self.app.setSize(800, 350)
        self.app.setResizable(canResize=False)
        self.app.setLocation("CENTER")
        self.app.setFont(size=11)
        # self.app.setButtonFont(size=12, family="Verdana")
        self.app.setStretch('both')

        # Frame 0
        self.app.startTabbedFrame("TabbedFrame")
        self.app.setTabbedFrameActiveBg('TabbedFrame', 'white')
        self.app.setTabbedFrameInactiveBg('TabbedFrame', '#F0F0F0')

        # Settings Tab
        self.app.startTab("Settings")
        self.app.setPadding([10, 0])
        self.app.addDirectoryEntry("input_dir", row=0, column=0,
                                   colspan=3).theButton.config(text="Input",
                                                               width=6)
        self.app.addCheckBox("Include Subdirs", row=0, column=3, colspan=1)
        self.app.setCheckBoxSticky('Include Subdirs', 'left')
        self.app.setCheckBox("Include Subdirs",
                             ticked=True,
                             callFunction=False)
        self.app.addDirectoryEntry("output_dir", row=1, column=0,
                                   colspan=3).theButton.config(text="Output",
                                                               width=6)
        self.app.setEntryChangeFunction('input_dir', self.on_input_dir_change)
        self.app.addLabelSpinBox("DPI",
                                 list(range(1200, 200, -100)),
                                 row=1,
                                 column=3,
                                 colspan=1)
        self.app.setSpinBox('DPI', 600, callFunction=False)
        self.app.setSpinBoxWidth('DPI', 8)
        self.app.setSpinBoxSticky('DPI', 'left')

        # H seperator
        self.app.addHorizontalSeparator(row=2,
                                        column=0,
                                        colspan=4,
                                        colour=None)

        self.app.addButton("Run", self.run, row=3, column=2)
        self.app.setButtonWidth('Run', 10)
        self.app.addCheckBox("Statistics", row=3, column=3)
        self.app.setCheckBox("Statistics", ticked=True, callFunction=False)

        self.app.stopTab()

        # Results Tab
        self.app.startTab("Results")
        self.app.startScrollPane("pane")
        self.app.stopScrollPane()
        self.app.stopTab()
        # Loggs Tab
        self.app.startTab("Logs")
        self.app.addScrolledTextArea("Logs")
        self.app.stopTab()
        self.app.stopTabbedFrame()

    def start_gui(self):
        """Start the frontend GUI"""
        logger.info('Starting frontend GUI')
        self.draw_app()
        self.app.go()

    def stop_gui(self):
        """Stop the frontend GUI"""
        if self.app.alive:
            logger.info('Stopping frontend GUI')
            self.app.stop()


def add_poppler_to_os_path(poppler_bin=None):
    """Add poppler binary path into OS environment 'PATH'.
    if poppler_bin not specified, `./poppler_bin` is used.
    """
    if poppler_bin is None:
        poppler_home = './poppler-0.68.0/bin'
    else:
        poppler_home = poppler_bin
    if os.path.isdir(poppler_home) is True:
        os.environ["PATH"] = poppler_home + os.pathsep + os.environ["PATH"]
    else:
        logger.warning('poppler executations do not exist at %r', poppler_home)


def gui_main():
    """The entry function to run the GUI of this programe
    """
    add_poppler_to_os_path()
    try:
        printer = MPrinterGUI()
        printer.start_gui()
    except Exception:
        printer.stop_gui()
        raise
