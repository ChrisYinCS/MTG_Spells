# -*- coding: utf-8 -*-
import SV_Data_To_Excel as SvExcel
import urllib2
import re

"""******************* Agent parameters for url ************************"""
url = "https://zh.moegirl.org/zh/%E4%B8%87%E6%99%BA%E7%89%8C/%E5%BC%82%E8%83%BD"
user_agent = ", like Gecko) Chrome/60.0.3112.113 Safari/537.36"
referer = "https://www.baidu.com/link?url=J-QZeQVLfvfZh7_lh8Qf0Vtk0ZsFxqsOCyw\
                GP4Ii3Howi8n5Ui5L6SvWPKS2OXiOG6ispEciqFLehmgDuUQudKnczccWhyDnUlP6TYWnKoYKzF_Ce\
                I_F6rvvdYVAqARG&wd=&eqid=849cd84b000029310000000659a6221b"

"""******************* Local file, used as a replacement when fail to access the website ************************"""
# local_file_addr = 'C:\Users\Administrator\PycharmProjects\Scrawler\MTG_Spells.txt'


def get_url(url_address):
    """
    get url using urllib module.
    :param url_address:
    :return:
    """
    req = urllib2.Request(url_address)
    req.add_header("user-agent", user_agent)
    req.add_header("referer", referer)
    req.add_header("GET", url_address)

    html = urllib2.urlopen(req)
    return html
    # try:
        # html = urllib2.urlopen(req)
        # return html
    # except urllib2.URLError, e:
    #     error_status=e.reason
    #     return error_status


def find_info(pattern, resource):
    """
    find the patterns in resource
    :param pattern: pattern to be search in the resource
    :param resource: resource to find the pattern in
    :return:the finditer result, it is a list of all matched results
    """
    result_list = re.finditer(pattern, resource)

    return result_list


def main():
    """
    main function of the shell
    :return: NULL
    """
    line_CH_names = 1       # the first line for Chinese names
    line_EN_names = 2       # the second line for English names
    line_Spell_type = 3     # the third line for spell types
    line_spell_details = 4  # the fourth line for spell details

    try:
        html_info_url = get_url(url).read()
    except urllib2.URLError, e:
        error_status=e.reason
        return error_status

    # f = open(local_file_addr, 'rb')
    # html_info_file = f.read()

    ExcelFile = SvExcel.ExcelOperate()

    '''******************** spell names ********************'''
    patterns_names = re.compile(r'<ul><li>( ?)<b>(.+?)</b>( ?)</li></ul>')  # 异能名称
    result_names = find_info(patterns_names, html_info_url)

    i = 2
    for each_item in result_names:
        changed_string = re.sub(r'</?\w+>', '', each_item.group())      # 异能名称

        Name_CH = re.search(r'^[^(\xef)]+', changed_string).group()     # 中文名称
        ExcelFile.SaveToExcel(i, line_CH_names, Name_CH.decode("utf-8").encode("gbk"))

        Name_EN = re.search(r'(\w ?\'?)+',changed_string).group()       # 英文名称
        ExcelFile.SaveToExcel(i, line_EN_names, Name_EN.decode("utf-8").encode("gbk"))
        i += 1

    """****************** spell type and details ***************"""
    pattern_details = re.compile(r'<dl><dd>((.+?(\n)?)+?)</dd></dl>')  # 异能描述
    result_details = find_info(pattern_details, html_info_url)

    i = 2
    for each_item in result_details:
        changed_string = re.sub(r'</?\w+>', '', each_item.group())
        # print ("%s: %s" % (i, repr(changed_string)))

        match_result = re.search(r'^.{6}\xe5\xbc\x8f\xe5\xbc\x82\xe8\x83\xbd', changed_string)  # 异能类型
        if match_result is not None:
            spell_type = match_result.group()
        else: spell_type = '/'
        # print spell_type
        ExcelFile.SaveToExcel(i, line_Spell_type, spell_type.decode("utf-8").encode("gbk"))

        spell_detail = re.sub(r'^.{6}\xe5\xbc\x8f\xe5\xbc\x82\xe8\x83\xbd\xe3\x80\x82', '', changed_string) # 异能描述
        # print spell_detail
        ExcelFile.SaveToExcel(i, line_spell_details, spell_detail.decode("utf-8").encode("gbk"))
        i += 1


if __name__=='__main__':
    main()
