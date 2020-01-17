#!/usr/bin/env python
# -*- coding: utf-8 -*-
""" The backend functions.
Author: sccotte@gmail.com
"""

import logging
import os
import platform
import re
from glob import glob
from os.path import basename, join
from subprocess import PIPE, Popen
from tempfile import NamedTemporaryFile

import pandas as pd
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image

logger = logging.getLogger('Backend')
logger.setLevel(logging.DEBUG)


class MinvoiceException(Exception):
    """Exception class for all backend functions.
    """
    def __init__(self, message):
        """Constructor"""
        self.message = message

    def __str__(self):
        """Reprentation of the instance"""
        return repr(self.message)


def find_pdfs(pdfs_dir, recursive=True, exclude_basenames=None):
    """ Find the invoice PDF file from the specified directory. The invoice
        PDF file should contains only one page.
    """
    if not pdfs_dir.endswith('/'):
        pdfs_dir += '/'
    # for windows, iglob is case-insensitive
    if recursive is True:
        pdfs = glob(pdfs_dir + '**/*.pdf', recursive=True)
    else:
        pdfs = glob(pdfs_dir + '*.pdf', recursive=False)
    if not pdfs:
        raise MinvoiceException('No invoice PDF file found!')
    pdfs_return = []
    for pdf_path in pdfs:
        if exclude_basenames is not None and \
                basename(pdf_path) in exclude_basenames:
            continue
        page_count = pdfinfo_from_path(pdf_path)["Pages"]
        if page_count != 1:
            msg = """Fatal error in processing PDF file {}: the invoice should
be one page per file""".format(pdf_path)
            raise MinvoiceException(msg)
        pdfs_return.append(pdf_path)
    return pdfs_return


def _get_command_path(command, poppler_path=None):
    """Return poppler command path from program name.
    """
    if platform.system() == "Windows":
        command = command + ".exe"
    if poppler_path is not None:
        command = os.path.join(poppler_path, command)
    return command


def pdftotext_from_path(pdf_path, text_path, poppler_path=None):
    """ Extract all pdf text to a txt file.
        The external program `pdftotext` is used, with option -layout
    """
    try:
        command = [_get_command_path("pdftotext", poppler_path), '-layout']

        command.extend([pdf_path, text_path])
        # Add poppler path to LD_LIBRARY_PATH
        env = os.environ.copy()
        if poppler_path is not None:
            env["LD_LIBRARY_PATH"] = poppler_path + ":" + env.get(
                "LD_LIBRARY_PATH", "")
        proc = Popen(command, env=env, stdout=PIPE, stderr=PIPE)
        out, err = proc.communicate()
        logger.debug('out=%s, error=%s', out, err)
    except (OSError, ValueError) as err:
        logger.error(err)
        raise


def to_text_str(pdf_path, poppler_path=None):
    """Extract text from a PDF file and return it as string.
    """
    ntp = NamedTemporaryFile('w+t', suffix='.txt', delete=False)
    ntp.close()
    pdftotext_from_path(pdf_path, ntp.name, poppler_path=poppler_path)
    with open(ntp.name, 'rt', encoding='utf-8') as fp:
        text_str = fp.read()
    os.unlink(ntp.name)
    return text_str


def parse_text(text_str):
    """ Parse the text content from invoice PDF to get name, phone and billing
        amount from a user.
    """
    user_name_regex = r'名\s*称\W\s?([\u4e00-\u9fff]{2,4})\s+'
    user_phoneno_regex = r'号码\W\s?(\d{11})'
    bill_date_regex = r'账期\W\s?(\d{6})'
    bill_amount_regex = r'\W小写\W+(\d+\.\d{2})'
    # user_name_regex = r'称(:|：)\s?([\u4e00-\u9fff]{2,4})\s+'
    # user_phoneno_regex = r'号码(:|：)\s?(\d{11})'
    # bill_date_regex = r'账期(:|：)\s?(\d{6})'
    # bill_amount_regex = r'(\(|）小写(\)).(\d+\.\d+)'
    r1 = re.search(user_name_regex, text_str, flags=re.U)
    r2 = re.search(user_phoneno_regex, text_str, flags=re.U)
    r3 = re.search(bill_date_regex, text_str, flags=re.U)
    r4 = re.search(bill_amount_regex, text_str, flags=re.U)
    any_fail_flag = False
    if r1:
        user_name = r1.group(1)
    else:
        any_fail_flag = True
        logger.error('Cannot extract user name with regular express %r',
                     user_name_regex)
    if r2:
        user_phoneno = r2.group(1)
    else:
        any_fail_flag = True
        logger.error(
            'Cannot extract user phone number with regular express %r',
            user_phoneno_regex)
    if r3:
        bill_date = r3.group(1)
    else:
        any_fail_flag = True
        logger.error(
            'Cannot extract user billing date with regular express %r',
            bill_date_regex)
    if r4:
        bill_amount = r4.group(1)
    else:
        any_fail_flag = True
        logger.error(
            'Cannot extract user billing amount with regular express %r',
            bill_amount_regex)
    if any_fail_flag is False:
        return (user_name, user_phoneno, bill_date, bill_amount)
    raise MinvoiceException(
        'Parse invoice text failed: {}, check logs'.format(text_str))


def pdf_to_jpg(pdfs, dpi=600):
    """Coverte pdf to jpe image file with specified dpi"""
    images = []
    for pdf_path in pdfs:
        # to image
        images += convert_from_path(pdf_path, dpi=dpi)
    return images


def to_raw_jpg_pdf(images, outfile):
    """ Save jpg images as a PDF file, each jpg image as one pdf page.
    """
    images[0].save(outfile, "PDF", save_all=True, append_images=images[1:])


def box_from_a4(image_size, a4_size, top_or_bottom=True):
    """Get the box when paste image into a4 sized page. If image size is
    overflow, then a resize operation on longer side would be taken while
    the with/height ratio is kept.

    Parameters:
        image_size: tuple
          (image_width, image_height)
        a4_size: tuple
          (a4_width, a4_height)
        top_or_bottom: boolean
          True: image is placed at top half of the a4 page
          False: image is place at bottom half of the a4 page
    """
    im_w, im_h = image_size
    a4_w, a4_h = a4_size

    w_r = im_w / a4_w
    h_r = im_h / a4_h * 2
    im_re_size = None
    # image overflow the a4, resize is needed
    if w_r > 1 or h_r > 1:
        # width is much larger
        if w_r > h_r:
            im_re_w, im_re_h = (a4_w, int(im_h / w_r))
            if top_or_bottom is True:
                box_top_left = (0, int((a4_h - im_re_h * 2) / 4))
            else:
                box_top_left = (0, int((a4_h - im_re_h * 2) / 2) + im_re_h)
        # height is much larger
        else:
            im_re_w, im_re_h = (int(im_w / h_r), int(a4_h / 2))
            if top_or_bottom is True:
                box_top_left = (int((a4_w - im_re_w) / 2), 0)
            else:
                box_top_left = (int((a4_w - im_re_w) / 2), im_re_h)
        im_re_size = (im_re_w, im_re_h)
    else:
        if top_or_bottom is True:
            box_top_left = (int((a4_w - im_w) / 2), int((a4_h - im_h) / 4))
        else:
            box_top_left = (int((a4_w - im_w) / 2), int(
                (a4_h - im_h) / 2) + im_h)
    return box_top_left, im_re_size


def place_one_jpg_on_a4_page(image, a4_page, top_or_bottom):
    """Place one jpg image on the A4 sized page.
    """
    box_top_left, im_re_size = box_from_a4(image.size,
                                           a4_page.size,
                                           top_or_bottom=top_or_bottom)
    if im_re_size is not None:
        resized_image = image.resize(im_re_size)
        a4_page.paste(resized_image, box=box_top_left)
    else:
        a4_page.paste(image, box=box_top_left)


def to_a4_jpg_pdf(images, outfile, dpi=600):
    """ Save jpg images as a A4-sized PDF file, two jpg images are placed
    in the upper and lower center of the A4 paper. Resize operation could
    be taken.
    """
    # A4 size in pixel
    a4_size = int(8.27 * dpi), int(11.7 * dpi)
    pages = []
    for i in range(0, len(images), 2):
        page = Image.new('RGB', a4_size, 'white')
        place_one_jpg_on_a4_page(images[i], page, top_or_bottom=True)
        if (i + 1) < len(images):
            place_one_jpg_on_a4_page(images[i + 1], page, top_or_bottom=False)
        pages.append(page)
    pages[0].save(outfile, "PDF", save_all=True, append_images=pages[1:])


def print_invoice(input_dir,
                  output_dir,
                  output_filenames,
                  recursive=True,
                  dpi=600,
                  do_analysis=True,
                  poppler_path=None):
    """The main backgroud function that read pdf files from specified folder,
    then do analysis, convert, and save the output.
    """
    outfile_statistics = join(output_dir, output_filenames['stats'])
    outfile_a4_jpg_pdf = join(output_dir, output_filenames['pdf'])
    logger.info('Finding PDF files in directory %r', input_dir)
    pdfs = find_pdfs(input_dir,
                     recursive=recursive,
                     exclude_basenames=output_filenames.values())
    logger.info('PDF files are: %r', pdfs)
    stats = []
    if do_analysis is True:
        logger.info('Searching text in the invoice PDF files ...')
        for pdf_path in pdfs:
            text_str = to_text_str(pdf_path, poppler_path=poppler_path)
            stats.append(parse_text(text_str) + (pdf_path, ))
        # (user_name, user_phoneno, bill_date, bill_amount)
        logger.info('Aggregating the data ...')
        df = pd.DataFrame(stats,
                          columns=[
                              'user_name', 'user_phoneno', 'bill_date',
                              'bill_amount', 'pdf_path'
                          ])
        df['user_name'] = df.user_name.astype('str')
        df['user_phoneno'] = df.user_phoneno.astype('str')
        df['bill_date_dt'] = pd.to_datetime(df.bill_date, format='%Y%m')
        df['bill_amount'] = df.bill_amount.astype('float')
        df['pdf_path'] = df.pdf_path.astype('str')
        df.sort_values(by=['user_name', 'user_phoneno', 'bill_date_dt'],
                       inplace=True)
        df['bill_year_quarter'] = df.bill_date_dt.dt.year.astype(
            'str') + '年第' + df.bill_date_dt.dt.quarter.astype('str') + '季度'
        quarter_df_per_phone = df.groupby(
            ['user_phoneno', 'bill_year_quarter']).bill_amount.sum().rename(
                'quarter_per_phone_per_user').reset_index()
        quarter_df_per_user = df.groupby([
            'user_name', 'bill_year_quarter'
        ]).bill_amount.sum().rename('quarter_per_user').reset_index()
        df = pd.merge(df,
                      quarter_df_per_phone,
                      on=['user_phoneno', 'bill_year_quarter'])
        df = pd.merge(df,
                      quarter_df_per_user,
                      on=['user_name', 'bill_year_quarter'])
        df.sort_values(
            by=['user_name', 'bill_year_quarter', 'user_phoneno', 'bill_date'],
            inplace=True)
        mindex_columns = [
            'user_name', 'bill_year_quarter', 'quarter_per_user',
            'user_phoneno', 'quarter_per_phone_per_user', 'bill_date'
        ]
        writen_columns = ['bill_amount']
        chinese_header_dict = dict(user_name=r'姓名',
                                   bill_year_quarter=r'年份季度',
                                   quarter_per_user=r'个人季度小计',
                                   user_phoneno=r'手机号',
                                   quarter_per_phone_per_user=r'个人每号码季度小计',
                                   bill_date=r'账单月',
                                   bill_amount=r'账单金额')
        df_excel = df[mindex_columns +
                      writen_columns].rename(columns=chinese_header_dict)
        df_excel = df_excel.set_index(
            [chinese_header_dict[c] for c in mindex_columns])
        logger.info('Saving statistics results as excel: %r',
                    outfile_statistics)
        df_excel.to_excel(
            outfile_statistics,
            sheet_name=r'高性能计算应用中心季度通讯费',
            columns=[chinese_header_dict[c] for c in writen_columns])
        pdfs = df.pdf_path.tolist()
    logger.info('Converting PDF to jpg ...')
    images = pdf_to_jpg(pdfs, dpi=dpi)
    # to_raw_jpg_pdf(images, outfile_raw_jpg_pdf)
    logger.info(
        'Saving images as PDF file with two images on A4-sized page %r',
        outfile_a4_jpg_pdf)
    to_a4_jpg_pdf(images, outfile_a4_jpg_pdf, dpi=dpi)
    logger.info('Backend Done')
    return df_excel
