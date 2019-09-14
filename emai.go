package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"gopkg.in/gomail.v2"
	"html/template"
	"io/ioutil"
	"regexp"
	"strconv"
	"strings"
	"time"
)

var (
	Emailconfig EmailConfig
	Judge       string = "YES"
	Sheetname   string
	Entry       []string
	Enum        int
)

type EmailConfig struct {
	Userinfo struct {
		User string `json:"user"`
		Pass string `json:"pass"`
		Host string `json:"host"`
		Port string `json:"port"`
	}
	Mailinfo struct {
		Title string `json:"title"`
		Pre   string `json:"pre"`
		Suf   string `json:"suf"`
	}
	Xlsinfo struct {
		Xls     string `json:"xls"`
		Sheetid string `json:"sheetid"`
	}
}

func main() {
	readConfig()
	xlsfile := readXls()
	echoInfo()
	sendMail(xlsfile)

}

func echoInfo() {
	fmt.Println("表格文件名：", Emailconfig.Xlsinfo.Xls)
	fmt.Println("sheet名：", Sheetname)
	fmt.Println("发送邮箱：", Emailconfig.Userinfo.User)
	fmt.Println("邮件主题：", Emailconfig.Mailinfo.Title)
	fmt.Println("总计有 ", Enum, " 封邮件需要发送！")
	fmt.Println("请确认是否发送邮件（YES/NO）:")
	fmt.Scanln(&Judge)
	if strings.ToUpper(Judge) == "YES" {
		fmt.Println("邮件发送开始 ！")
	} else if strings.ToUpper(Judge) == "Y" {
		fmt.Println("邮件发送开始 ！")
	}else {
		panic("取消发送")
	}

}

func readConfig() {
	data, err := ioutil.ReadFile("./email.config")
	checkErr(err)
	err = json.Unmarshal(data, &Emailconfig)
	checkErr(err)
}

func sendMail(rows [][]string) {
	var touser string
	for _, row := range rows[1:] {
		buffer := new(bytes.Buffer)
		matched, err := regexp.MatchString("\\w[-\\w.+]*@([A-Za-z0-9][-A-Za-z0-9]+\\.)+[A-Za-z]{2,14}", row[len(row)-1])
		if matched {
			touser = row[len(row)-1]

		} else {
			fmt.Println("获取邮箱失败,错误的邮箱格式,姓名：", row[0], "邮箱信息", row[len(row)-1])
			continue
		}
		checkErr(err)
		t, _ := template.ParseFiles("./email_template.html")
		t.Execute(buffer, struct {
			SEntry []string
			Srow   []string
			Spre   string
			Ssuf   string
			Stime  string
		}{
			SEntry: Entry,
			Srow:   row,
			Spre:   Emailconfig.Mailinfo.Pre,
			Ssuf:   Emailconfig.Mailinfo.Suf,
			Stime:  time.Now().Format("2006年01月02日"),
		})
		nmail := 0
		nmail += 1
		fmt.Println("总共[",Enum,"]封邮件,开始发送第[",nmail,"]封邮件,发送给",touser)
		//send(touser, buffer.String())
		fmt.Println(touser)
	}
	fmt.Println("总共[",Enum,"]封邮件,现已全部发送完毕 !")
	fmt.Printf("\t")
	fmt.Println("输入任意字符退出程序 ！")
	fmt.Scanln(&Judge)

}

func send(touser string, body string) {
	m := gomail.NewMessage()
	m.SetHeader("From", Emailconfig.Userinfo.User)
	m.SetHeader("To", touser)
	//m.SetAddressHeader("Cc", "dan@example.com", "Dan")
	m.SetHeader("Subject", Emailconfig.Mailinfo.Title)
	m.SetBody("text/html", body)
	//m.Attach("/home/Alex/lolcat.jpg")

	iport, err := strconv.Atoi(Emailconfig.Userinfo.Port)
	checkErr(err)
	d := gomail.NewDialer(Emailconfig.Userinfo.Host, iport, Emailconfig.Userinfo.User, Emailconfig.Userinfo.Pass)

	// Send the email to Bob, Cora and Dan.
	if err := d.DialAndSend(m); err != nil {
		fmt.Println(err)
	}else {
		fmt.Println("邮件发送成功 ！")
	}

}

func readXls() (xlsfile [][]string) {
	xlsFile, err := excelize.OpenFile(Emailconfig.Xlsinfo.Xls)
	checkErr(err)
	isheetid, _ := strconv.Atoi(Emailconfig.Xlsinfo.Sheetid)
	Sheetname = xlsFile.GetSheetName(isheetid)
	rows, err := xlsFile.GetRows(Sheetname)
	Entry = rows[0]
	Enum = len(rows)-1
	return rows
}

func checkErr(err error) {
	if err != nil {
		fmt.Println(err)
	}
}
