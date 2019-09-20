package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"gopkg.in/gomail.v2"
	"html/template"
	"io/ioutil"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"
)

var (
	Emailconfig EmailConfig
	Judge       string = "YES"
	Sheetname   string
	Entry       [][]string
	Enum        int
	Nmail       int = 0
	Head        int = 1
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
		Head    string `json:"head"`
	}
}

type Sendinfo struct {
	Sbody	template.HTML
	Spre   string
	Ssuf   string
	Stime  string
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
	} else {
		os.Exit(0)
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

	for _, row := range rows[Head:] {
		buffer := new(bytes.Buffer)
		matched, err := regexp.MatchString("\\w[-\\w.+]*@([A-Za-z0-9][-A-Za-z0-9]+\\.)+[A-Za-z]{2,14}", row[len(row)-1])
		if matched {
			touser = row[len(row)-1]

		} else {
			fmt.Println("获取邮箱失败,错误的邮箱格式,姓名：", row[0], "邮箱信息", row[len(row)-1])
			continue
		}
		checkErr(err)
		body:=template.HTML(Emailbody(Entry,row))

		Sendinfo := Sendinfo{
			Sbody:	body,
			Spre:   Emailconfig.Mailinfo.Pre,
			Ssuf:   Emailconfig.Mailinfo.Suf,
			Stime:  time.Now().Format("2006年01月02日"),
		}

		t := template.New("email_template.html")
		t, _ = template.ParseFiles("email_template.html")
		t.Execute(buffer,Sendinfo)
		Nmail += 1
		fmt.Println("总共[", Enum, "]封邮件,开始发送第[", Nmail, "]封邮件,发送给", touser)
		//send(touser, buffer.String())
		fmt.Println(buffer.String())
		fmt.Println(touser)
	}
	fmt.Println("总共[", Enum, "]封邮件,已全部发送完毕 !")
	fmt.Printf("\t")
	fmt.Println("输入回车字符退出程序 ！")
	fmt.Scanln(&Judge)
	os.Exit(0)
}

func Emailbody(entry [][]string,row []string)(str string) {
	lnum:=len(entry)-1
	rnum:=len(entry[0])

	//	var str string
	for i:=0;i<rnum ;i++  {
		for j:=0;j<lnum ;j++  {
			var rspan,cspan int = 1,1
			// i=23 j=1
			if entry[j][i] != "" {
				for z:=i;z<rnum-1 ;z++  {
					if entry[j][z+1] =="" {
						rspan += 1
					}else {
						break
					}
				}
			}

			if j+1<lnum {
				if entry[j+1][i] == "" {
					cspan += 1
				}
			}
			srspan := strconv.Itoa(rspan)
			scspan := strconv.Itoa(cspan)
			if entry[j][i] != "" {
				str += "<th rowspan=\"" + srspan + "\"colspan=\"" + scspan + "\">" + entry[j][i] + "</th>"
			}
		}
		str=str+"<td>" + row[i] + "</td>"
		str="<tr>"+str+"</tr>"
	}
	return str
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
	} else {
		fmt.Println("邮件发送成功 ！")
	}

}

func readXls() (xlsfile [][]string) {
	xlsFile, err := excelize.OpenFile(Emailconfig.Xlsinfo.Xls)
	checkErr(err)
	isheetid, _ := strconv.Atoi(Emailconfig.Xlsinfo.Sheetid)
	Sheetname = xlsFile.GetSheetName(isheetid)
	rows, err := xlsFile.GetRows(Sheetname)
	Head, _ = strconv.Atoi(Emailconfig.Xlsinfo.Head)
	Entry = rows[0:Head]
	Enum = len(rows) - Head
	return rows
}

func checkErr(err error) {
	if err != nil {
		fmt.Println(err)
	}
}
