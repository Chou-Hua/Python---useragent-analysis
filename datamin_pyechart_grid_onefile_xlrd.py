# coding=utf-8
from pyecharts.charts import Bar, Grid, Page
from pyecharts import options as opts
import os
from user_agents import parse
import randomcolor
import pathlib
import xlrd



def datacompress(dictname, isrank, sheetname1):
    PC = Bar()
    iOS = Bar()
    Android = Bar()
    listname = ['iOS', 'Android', 'PC']
    version = ['Version']
    barlist = [iOS, Android, PC]
    for i in range(len(dictname)):
        sortdict = {}
        for x in dictname[i]:
            sortdict[x] = (sum(dictname[i][x]))
        sortdict = dict(
            (sorted(sortdict.items(), key=lambda x: x[1])))  # 轉轉轉排序再轉轉轉
        sortdictkey = (list(sortdict.keys()))
        sortdictkey.reverse()
        sortdictvalue = (list(sortdict.values()))
        sortdictvalue.reverse()
        if isrank == True:
            colors = randomcolor.RandomColor().generate()
            barlist[i] = (Bar()
                          .add_xaxis(sortdictkey[0:5])
                          .add_yaxis(version[0], (sortdictvalue[0:5]), color=colors, category_gap="50%")
                          )            
        else:
            barlist[i] = (Bar()
                          .add_xaxis(list(sortdict.keys()))
                          .add_yaxis(listname[i], list(sortdict.values()), color=randomcolor.RandomColor().generate(), category_gap="50%")
                          )
    return barlist


def board(duseagntlist, oslist, sheetname, xlsxname):
    page = Page(interval=1)
    data = datacompress(duseagntlist, False, sheetname)
    osdata = datacompress(oslist, True, sheetname)
    #print(data[0])
    iOS = data[0].set_global_opts(legend_opts=opts.LegendOpts(pos_left="20%"),
                                  title_opts=opts.TitleOpts(title="iOS Browser Rank(全)"), xaxis_opts=opts.AxisOpts(name="瀏覽器名稱", axislabel_opts={"rotate": "30", "interval": "0"})
                                  )
    Android = data[1].set_global_opts(
        legend_opts=opts.LegendOpts(pos_left="20%"),
        title_opts=opts.TitleOpts(title="Android Browser Rank(全)"), xaxis_opts=opts.AxisOpts(name="瀏覽器名稱", axislabel_opts={"rotate": "30", "interval": "0"}),
    )
    PC = data[2].set_global_opts(legend_opts=opts.LegendOpts(pos_left="20%"),
                                 title_opts=opts.TitleOpts(title="PC Browser Rank(全)"), xaxis_opts=opts.AxisOpts(name="瀏覽器名稱", axislabel_opts={"rotate": "30", "interval": "0"})
                                 )
    osios = osdata[0].set_global_opts(
        legend_opts=opts.LegendOpts(pos_right="20%"),
        title_opts=opts.TitleOpts(title="iOS version Rank(前五)", pos_right="5%"), xaxis_opts=opts.AxisOpts(name="系統版本", axislabel_opts={"interval": "0"})
    )
    osand = osdata[1].set_global_opts(
        legend_opts=opts.LegendOpts(pos_right="20%"),
        title_opts=opts.TitleOpts(title="Android version Rank(前五)", pos_right="5%"), xaxis_opts=opts.AxisOpts(name="系統版本", axislabel_opts={"interval": "0"}),
    )
    ospc = osdata[2].set_global_opts(
        legend_opts=opts.LegendOpts(pos_right="20%"),
        title_opts=opts.TitleOpts(title="PC version Rank(前五)", pos_right="5%"), xaxis_opts=opts.AxisOpts(name="系統版本", axislabel_opts={"interval": "0"})
    )
    iosgrid = (Grid(init_opts=opts.InitOpts(width="1680px", height="500px"))
               .add(iOS, grid_opts=opts.GridOpts(pos_right="55%", pos_bottom="20%"))
               .add(osios, grid_opts=opts.GridOpts(pos_left="55%", pos_bottom="20%"))
               )
    aosgrid = (Grid(init_opts=opts.InitOpts(width="1680px", height="500px"))
               .add(Android, grid_opts=opts.GridOpts(pos_right="55%", pos_bottom="20%"))
               .add(osand, grid_opts=opts.GridOpts(pos_left="55%", pos_bottom="20%"))
               )
    pcosgrid = (Grid(init_opts=opts.InitOpts(width="1680px", height="500px"))
                .add(PC, grid_opts=opts.GridOpts(pos_right="55%", pos_bottom="20%"))
                .add(ospc, grid_opts=opts.GridOpts(pos_left="55%", pos_bottom="20%"))
                )
    page.add(iosgrid)
    page.add(aosgrid)
    page.add(pcosgrid)
    page.render(xlsxname +"_ "+ sheetname + "_" + ".html")


path = os.getcwd()
filename = r"*.xlsx" #所有xlsx檔案
xlsxlist = list(pathlib.Path(path).glob(filename))
for xl in range(len(xlsxlist)):
    name = (xlsxlist[xl].name).split(".")
    xls = xlrd.open_workbook(xlsxlist[xl])        
    sheetlist = xls.sheet_names()    
    for i in range(len(sheetlist)):
        mobile = []
        sheetname = xls.sheet_by_name(sheetlist[i])
        Amobilecount = {}
        imobilecount = {}
        pccount = {}
        Aos = {}
        ipos = {}
        pcos = {}
        phonetcoutlist = []
        oscoutlist = []                
        size1 = (sheetname.nrows)-1            
        for k in range(size1):
            user_string = (sheetname.row_values(k+1))                    
            user_agent = parse(str(user_string))
            if ((user_agent.is_mobile == True and user_agent.os.family == 'Android') or (user_agent.is_tablet == True and user_agent.os.family == 'Android')):
                if user_agent.os.family + " " + user_agent.os.version_string not in Aos.keys():
                    Aos[user_agent.os.family + " " +
                        user_agent.os.version_string] = [int(float(sheetname.cell_value(k+1,1)))]
                else:
                    Aos[user_agent.os.family + " " +
                        user_agent.os.version_string].append((int(float(sheetname.cell_value(k+1,1)))))
                if user_agent.browser.family not in Amobilecount.keys():
                    Amobilecount[user_agent.browser.family] = [int(float(sheetname.cell_value(k+1,1)))]
                else:
                    Amobilecount[user_agent.browser.family].append(
                        int(float(sheetname.cell_value(k+1,1))))
            elif((user_agent.is_mobile == True and user_agent.os.family == 'iOS') or (user_agent.is_tablet == True and user_agent.os.family == 'iOS')):
                if user_agent.os.family + " " + user_agent.os.version_string not in ipos.keys():
                    ipos[user_agent.os.family + " " +
                         user_agent.os.version_string] = [int(float(sheetname.cell_value(k+1,1)))]
                else:
                    ipos[user_agent.os.family + " " +
                         user_agent.os.version_string].append(int(float(sheetname.cell_value(k+1,1))))
                if user_agent.browser.family not in imobilecount.keys():
                    imobilecount[user_agent.browser.family] = [int(float(sheetname.cell_value(k+1,1)))]
                else:
                    imobilecount[user_agent.browser.family].append(
                        int(float(sheetname.cell_value(k+1,1))))
            else:
                if user_agent.os.family + " " + user_agent.os.version_string not in pcos.keys():
                    pcos[user_agent.os.family + " " +
                         user_agent.os.version_string] = [int(float(sheetname.cell_value(k+1,1)))]
                else:
                    pcos[user_agent.os.family + " " +
                         user_agent.os.version_string].append(int(float(sheetname.cell_value(k+1,1))))
                if user_agent.browser.family not in pccount.keys():
                    pccount[user_agent.browser.family] = [int(float(sheetname.cell_value(k+1,1)))]
                else:
                    pccount[user_agent.browser.family].append(int(float(sheetname.cell_value(k+1,1))))
        oscoutlist.append(ipos)
        oscoutlist.append(Aos)
        oscoutlist.append(pcos)
        phonetcoutlist.append(imobilecount)
        phonetcoutlist.append(Amobilecount)
        phonetcoutlist.append(pccount)
        board(phonetcoutlist, oscoutlist, sheetlist[i],name[0])
