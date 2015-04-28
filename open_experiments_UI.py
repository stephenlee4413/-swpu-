#-*- coding:utf-8 -*-
##############################
### Author:lee          ######
### create_date:2015-04-03 ###
### copyright: lee_SWPU    ###
##############################

import ttk
from Tkinter import *
from read_excel import *

root = Tk()
root.title("开放性实验项目")


###规划整个UI界面的布局问题，使用Grid进行布局####
content = ttk.Frame(root)
main_frame = ttk.Frame(content, borderwidth=5, relief="sunken", width=400, height=200)


###software use tips###
notice = """
		软件使用申明:
	本项目受西南石油大学开放性实验项目资助，
	主要用于教学使用，不得用于其他商业目的，
	特此说明。
					现代教育技术中心
"""
# notice_label = ttk.Label(main_frame, text=notice, foreground="red").grid(column=5, row=0, sticky=N)
notice_label = ttk.Label(main_frame, text=notice, foreground="red")


####将从excel表格中读取的信息显示到label中#######
info_show = info_show_from_excel() ##excel中的数据信息
print info_show
info_show_label = ttk.Label(main_frame, width=70, text="以下内容显示本次等考考务相关信息:")
info_show_text = Text(main_frame, width=70, height=15)
info_show_text.insert(INSERT, info_show)

#####grid布局控制################################
content.grid(column=0, row=0, sticky=(N, S, E, W))
main_frame.grid(row=0, column=0)
notice_label.grid(row=1, column=2, columnspan=3, sticky=N, padx=5)
info_show_label.grid(column=2, row=2, columnspan=3, sticky=N)
info_show_text.grid(column=2, row=3, columnspan=3, sticky=N)

root.mainloop()