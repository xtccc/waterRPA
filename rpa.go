// 原作者水哥
package main

import (
	"fmt"
	"runtime"
	"strconv"
	"time"

	"github.com/atotto/clipboard"
	"github.com/extrame/xls"
	"github.com/go-vgo/robotgo"
	"github.com/kpango/glg"
	"github.com/micmonay/keybd_event"
)

func clear_bat() {
	// 清除 位于 %temp% 包含_MEI 的文件夹，
}
func ctrl_v() {
	kb, err := keybd_event.NewKeyBonding()
	if err != nil {
		panic(err)
	}

	// For linux, it is very important to wait 2 seconds
	if runtime.GOOS == "linux" {
		time.Sleep(2 * time.Second)
	}

	// Select keys to be pressed
	kb.SetKeys(keybd_event.VK_V)

	// Set shift to be pressed
	//kb.HasSHIFT()
	kb.HasCTRL(true)

	// Press the selected keys
	err = kb.Launching()
	if err != nil {
		panic(err)
	}
}

func click(clickTimes int, lOrR string, x int, y int) {
	robotgo.MouseSleep = 100
	robotgo.Move(x, y)
	if clickTimes == 2 {
		robotgo.Click(lOrR)
		robotgo.Click(lOrR)
	} else {
		robotgo.Click(lOrR)
	}

}
func mouseClick(clickTimes int, lOrR string, img string, reTry int) {
	if reTry == 1 {
		for {
			pp := get_position_from_img(img)
			if len(pp) > 0 {
				glg.Infof("Found %d matches:\n", len(pp))
				for _, p := range pp {
					glg.Debug("- (%d, %d) with %f accuracy\n", p.X, p.Y, p.G)
					click(clickTimes, lOrR, p.X, p.Y)
				}
				break
			} else {
				glg.Info("未找到匹配图片,0.1秒后重试")
				time.Sleep(time.Millisecond * 100)
			}

		}
	} else if reTry == -1 { // -1 永远重复
		for {
			pp := get_position_from_img(img)
			if len(pp) > 0 {
				glg.Info("Found %d matches:\n", len(pp))
				for _, p := range pp {
					glg.Debug("- (%d, %d) with %f accuracy\n", p.X, p.Y, p.G)
					click(clickTimes, lOrR, p.X, p.Y)
				}
				break
			} else {
				glg.Info("未找到匹配图片,0.1秒后重试")
				time.Sleep(time.Millisecond * 100)
			}
		}
	} else if reTry > 1 {
		i := 1
		for i < reTry+1 {
			pp := get_position_from_img(img)
			if len(pp) > 0 {
				glg.Info("Found %d matches:\n", len(pp))
				for _, p := range pp {
					glg.Debug("- (%d, %d) with %f accuracy\n", p.X, p.Y, p.G)
					click(clickTimes, lOrR, p.X, p.Y)
				}
				break
			} else {
				glg.Info("未找到匹配图片,0.1秒后重试")
				time.Sleep(time.Millisecond * 100)
			}
		}
	}
}
func mainWork(sheet1 *xls.WorkSheet) {
	for i := 1; i <= int(sheet1.MaxRow); i++ { // 遍历rows ,遍历每一行
		glg.Debug("sheet1.MaxRow", sheet1.MaxRow)
		glg.Debug("i", i)
		//取本行指令的操作类型
		col1_str := sheet1.Row(i).Col(0)
		glg.Debug("col1_str", col1_str)
		if len(col1_str) == 0 {
			break
		}
		//string to int
		col1_int, err := strconv.Atoi(col1_str)
		if err != nil {
			glg.Error(err)
		}
		switch col1_int {
		case 1, 2, 3:
			//取图片名称
			img := sheet1.Row(i).Col(1)
			reTry := 1
			//第三列的类型是数字，且第三列不为0
			if col3_int, err := strconv.Atoi(sheet1.Row(i).Col(2)); err == nil && col3_int != 0 {
				reTry, _ = strconv.Atoi(sheet1.Row(i).Col(2)) //获取重试次数
			}
			if col1_int == 3 {
				mouseClick(1, "right", img, reTry)
				glg.Info("右键", img)
			} else {
				glg.Info(fmt.Sprintf("%d次击打左键", col1_int), img)
				mouseClick(col1_int, "left", img, reTry)
			}
		case 4: //4代表输入文本
			inputValue := sheet1.Row(i).Col(1)
			clipboard.WriteAll(inputValue) //把文本写入剪贴板
			ctrl_v()
			glg.Info("ctrl+v:", inputValue)
			time.Sleep(time.Millisecond * 500)

		case 5: //5代表等待
			//取图片名称
			waitTime, _ := strconv.Atoi(sheet1.Row(i).Col(1))
			glg.Info("等待", waitTime, "秒")
			time.Sleep(time.Millisecond * 1000 * time.Duration(waitTime))
		case 6: //6代表滚轮
			//取图片名称
			scroll, _ := strconv.Atoi(sheet1.Row(i).Col(1))
			robotgo.MouseSleep = 100
			glg.Debug("滚轮滑动", int(scroll), "距离")
			robotgo.Scroll(0, scroll)

		}
	}
}
func dataCheck(sheet1 *xls.WorkSheet) bool {
	// 数据检查
	// cmdString.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
	// ctype     空：0
	//           字符串：1
	//           数字：2
	//           日期：3
	//           布尔：4
	//           error：5
	checkCmd := true
	//行数检查
	if sheet1.MaxRow < 1 {
		glg.Error("没数据啊哥")
		checkCmd = false
	}
	//每行数据检查

	for i := 1; i <= int(sheet1.MaxRow); i++ { // 遍历rows ,遍历每一行
		glg.Debug("sheet1.MaxRow", sheet1.MaxRow)
		glg.Debug("i", i)
		// 第1列 操作类型检查
		col1_str := sheet1.Row(i).Col(0)
		glg.Debug("col1_str", col1_str)
		if len(col1_str) == 0 {
			break
		}
		// string to int
		col1_int, err := strconv.ParseInt(col1_str, 0, 64)
		if err != nil {
			glg.Error("第", i+1, "行", "第1列数据错误")
			checkCmd = false
		}
		// 第2列 内容检查
		col2_str := sheet1.Row(i).Col(1)
		// 读图点击类型指令，内容必须为字符串类型
		// 第一列是1，2，3 ,第二列值必须为*.png
		if col1_int == 1 || col1_int == 2 || col1_int == 3 {

			if col2_str[len(col2_str)-4:] != ".png" { // 如果不是 .png
				glg.Error(fmt.Sprintf("第%d行", i+1), "第2列数据错误")
				glg.Error("读图点击类型指令，内容必须为.png 结尾的字符串")
				glg.Error("内容是", col2_str)
				checkCmd = false
			}

		}
		// 输入类型，内容不能为空
		// 第一列是4,第二列不能为空
		if col1_int == 4 {
			if len(col2_str) <= 0 {
				print('第', i+1, "行,第2列数据有毛病")
				checkCmd = false
			}
		}
		// 等待类型，内容必须为数字
		// 第一列是5,第二列必须为数字
		if col1_int == 5 {
			_, err := strconv.ParseFloat(col2_str, 64)
			if err != nil {
				glg.Error("第", i+1, "行", "第2列数据错误")
				checkCmd = false
			}
		}

		// 滚轮事件，内容必须为数字
		//第一列是6,第二列必须为数字
		if col1_int == 6 {
			_, err := strconv.ParseFloat(col2_str, 64)
			if err != nil {
				glg.Error("第", i+1, "行", "第2列数据错误")
				checkCmd = false
			}
		}
	}
	return checkCmd
}
func main() {
	//glg 设置日志等级
	if runtime.GOOS == "windows" {
		glg.Get().DisableColor()
	}
	glg.Get().SetLevel(glg.INFO)
	//glg.Get().SetLevel(glg.DEBG)

	file := "cmd.xls"
	//打开文件
	//wb = xlrd.open_workbook(filename=file)
	xlFile, err := xls.Open(file, "utf-8")
	if err != nil {
		glg.Error(err)
	}
	// 获取表格sheeet页面
	sheet1 := xlFile.GetSheet(0)
	for i := 1; i <= int(sheet1.MaxRow); i++ {
		glg.Debug("sheet1.Row(", i, ").Col(0)", sheet1.Row(i).Col(0))
	}
	// 数据检查
	checkCmd := dataCheck(sheet1)
	glg.Debug("检查结果：", checkCmd)
	if checkCmd {
		glg.Info("选择功能: 1.做一次 2.循环到死 \n")
		var key string
		fmt.Scanln(&key)
		if key == "1" {
			//循环拿出每一行指令
			mainWork(sheet1)
		} else if key == "2" {
			for {
				mainWork(sheet1)
				//sleep 0.1s
				time.Sleep(time.Millisecond * 100)
				glg.Info("等待0.1秒")
			}
		}
	} else {
		glg.Info("输入有误或者已经退出!")
	}
	clear_bat()

}
