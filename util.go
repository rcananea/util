/*
** Arquivo: util.go by Cananéa
** Atualizado: 02 de Junho de 2022
*/

package util

import (
	"bufio"
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"runtime"
	"strings"
	"math/rand"
	"time"
	"bytes"
	"net"
	"database/sql"
	"github.com/tealeg/xlsx"
	_ "github.com/lib/pq"
)

// Constantes
const VERSAO = "0.23"
const AUTOR  = "Ronaldo Cananéa"
const Silver = "C0C0C0"
const White  = "FFFFFF"
const Yellow = "FFFFCC"
const Right  = "right"
const Center = "center"
const Left	 = "left"
const Pipe	 = "|"
const max    = 20
// Variável Global
var antes string = ""

type INFODB struct {
	Host string
	DBName string
	User string
	Password string
	sslmode string
	Port int
	Err error
}

func DateHour2Time(dh string) time.Time {
   DH,_ := time.Parse("2006-01-02 03:04:05",dh)
   return DH
}

func DiffTime(dh1 string,dh2 string) (df time.Duration) {
   t1 := DateHour2Time(dh1)
   t2 := DateHour2Time(dh2)
   df = t2.Sub(t1)
   return df
}

func ReadLines(path string) ([]string,int,error) {
	return readLines(path)
}

func ReadLinesSemAspas(path string) ([]string,error) {
	return readLinesSemAspas(path)
}

func CrieLOG(logname string) (*os.File,error,string) {
	namelog := filepath.Join(os.TempDir(),fmt.Sprintf("%s_%s",GetUserName(),logname))
	file,err := os.Create(namelog)
	return file,err,namelog
}

func WriteLines(lines [][]string, path string) error {
	return writeLines(lines,path)
}

func LExcel2Mem(pgm string,versao string,xlsFile string,show bool) (int,string) {
	var texto string = ""
	tmpFile := filepath.Join(os.TempDir(),fmt.Sprintf("_%s_%s_.txt",GetUserName(),GetFnameNoExtension(xlsFile)))
	erro,_ := LExcel(pgm,versao,xlsFile,tmpFile,show)
	if erro == 0 {
		buf,err := ioutil.ReadFile(tmpFile)
		if err != nil {
			erro = -99
		} else {
			texto = string(buf)
		}
	}
	_ = os.Remove(tmpFile)
	return erro,texto
}

func Agora() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s",ja[11:16])
}

func AddString2Mem(texto *string,linha string) {
	var buffer bytes.Buffer
	buffer.WriteString(*texto)
	buffer.WriteString(linha)
	*texto = buffer.String()
}

func FiltreMUL(csvin string,mul string)([][]string) {
	// csvin  - arquivo em formato csv gerado pelo IMM
	// mul    - nome da ligação ICCP ou I61850
	// inicio - true, se está no início da linha
	var rec [][]string
	lines,err := ProcessCSV(csvin,false)
	if err == nil {
		nr := 0
		ncampos := 0
		for ln := 0; ln < len(lines); ln++ {
			if ln == 0 {
				ncampos = len(lines[ln])
			}
			for p := 0; p < len(lines[ln]); p++ {
				if CompareSubString(lines[ln][p],mul,len(mul)) == 0 { // igual?
					nr += 1
					break
				}
			}
		}
		if nr == 0 {
			rec = make([][]string,1)
			var lnvazio []string
			for z := 0; z < ncampos; z++ {
				lnvazio = append(lnvazio," ")
			}
			rec[0] = lnvazio
		} else {
			rec = make([][]string,nr)
			nr = 0
			for ln := 0; ln < len(lines); ln++ {
				if ln > 0 {
					ok := false
					for p := 0; p < len(lines[ln]); p++ {
						if CompareSubString(lines[ln][p],mul,len(mul)) == 0 { // igual?
							ok = true
							break
						}
					}
					if ok {
						var lineout []string
						for p := 0; p < len(lines[ln]); p++ {
							lineout = append(lineout,lines[ln][p])
						}
						rec[nr] = lineout
						nr  += 1
					}
				}
			}
		}
	}
	return rec
}

// Se igual retorne 0
func CompareSubString(a,b string,length int) int {
	var ret int = 0
	lena := len(a)
	lenb := len(b)
	lenc := length
	if lenc > lena {
		lenc = lena
	}
	if lenc > lenb {
		lenc = lenb
	}
	sa := a[0:lenc]
	sb := b[0:lenc]
	if sa < sb {
		ret = -1
	} else if sa > sb {
		ret = 1
	}
	return ret
}

// Se igual retorne 0
func CompareString(a,b string) int {
	var ret int = 0
	if a < b {
		ret = -1
	} else if a > b {
		ret = 1
	}
	return ret
}

func GetContentsDir(dir string) ([]string,error) {
	var names []string
	d,err := os.Open(dir)
	if err == nil {
		defer d.Close()
		names,err = d.Readdirnames(-1)
	}
	return names,err
}

func GetFilesFromDir(dir string,mask string) ([]string,error) {
	var fnames []string
	names,err := GetContentsDir(dir)
	if err == nil {
		for n := 0; n < len(names); n++ {
			if strings.Contains(names[n],mask) {
				fn := filepath.Join(dir,names[n])
				file,erro := os.Open(fn)
				if erro == nil {
					defer file.Close()
					fi,er := file.Stat()
					if er == nil {
						if !fi.IsDir() {
							fnames = append(fnames,fn)
						} else {
							err = er
						}
					}
				}
			}
		}
	}
	return fnames,err
}

func RemoveAllCSVS(dir string) (int,int) {
	var notRemoved int = 0
	fnames,err := GetFilesFromDir(dir,".csv")
	if err == nil {
		for f := 0; f < len(fnames); f++ {
			if err = os.Remove(fnames[f]); err != nil {
				notRemoved++
			}
		}
	} else {
		notRemoved = len(fnames)
	}
	return notRemoved,len(fnames)
}

func EscreveLOG(w *bufio.Writer,msg string) {
	fmt.Fprintf(w,"%s\n",msg)
	w.Flush()
}

func GetFnameNoExtension(fname string) (string) {
	parts := strings.Split(GetFnameOnly(fname),".")
	return parts[0]
}

func Date2Time(s string) time.Time {
   d,_ := time.Parse("2006-01-02",s)
   return d
}

func CompareDates(DateA string,DateB string) (bool) {
   a := Date2Time(DateA)
   b := Date2Time(DateB)
   daysA := a.YearDay()
   for year := a.Year(); year < b.Year(); year++ {
      daysA += time.Date(year,time.December,31,0,0,0,0,time.UTC).YearDay()
   }
   daysB := b.YearDay()
   fmt.Println(daysA,daysB)
   return daysA <= daysB
}

func DaysBetween(a,b time.Time) int {
   if a.After(b) {
      a,b = b,a
   }
   days := -a.YearDay()
   for year := a.Year(); year < b.Year(); year++ {
      days += time.Date(year,time.December,31,0,0,0,0,time.UTC).YearDay()
   }
   days += b.YearDay()
   return days
}

func GetDaysBetweenDates(DateA string,DateB string) (int) {
   // DateA e DateB CCYY-MM-DD
   return DaysBetween(Date2Time(DateA),Date2Time(DateB))
}

func GetDateDaysAgo(nday int) (error,string) {
   var data string = ""
   var err error = nil
   if nday > 0 {
      nhr := time.Duration(nday * 24)
      dt := strings.Split(fmt.Sprintf("Yesterday: %v",time.Now().Add(-nhr*time.Hour))," ")
      data = dt[1]
   } else {
      err = fmt.Errorf("Esperando numero de dias maior do que 0: econtrado %d",nday)
   }
   return err,data
}

func GetDateYesterday() (error,string) {
   return GetDateDaysAgo(1)
}

func GetDateWeekAgo() (error,string) {
   return GetDateDaysAgo(7)
}

func GetDateMonthAgo() (error,string) {
   return GetDateDaysAgo(30)
}

func GetFileInfo(fn string) (int64,string) {
	var tam int64 = 0
	var data string = ""
	fi,err := os.Stat(fn)
	if err == nil {
		tam = fi.Size()
		fd := fi.ModTime()
		data = fmt.Sprintf("%02d-%02d-%04d %02d:%02d:%02d",fd.Day(),fd.Month(),fd.Year(),fd.Hour(),fd.Minute(),fd.Second())
	}
	return tam,data
}

func ToBR(tUTC time.Time)(time.Time) {
	secs := tUTC.Unix()
	return time.Unix(secs,0)
}

func ToFloat64(campo string) (float64,error) {
	x,err := strconv.ParseFloat(campo,64)
	if err != nil {
		var y int
		y,err = strconv.Atoi(campo)
		if err == nil {
			x = float64(y)
		}
	}
	return x,err
}

func GetWD() (string) {
	pwd,err := os.Getwd()
	if err != nil {
		pwd = ""
	}
	return pwd
}

func TimeZero()(time.Time) {
	return time.Unix(0,0)
}

func HeaderOnly(csvin string) ([][]string) {
	// csvin - arquivo em formato csv gerado pelo IMM
	rec := make([][]string,2)
	lines,err := ProcessCSV(csvin,false)
	if err == nil {
		var line []string
		nc := 0
		for ln := 0; ln < len(lines[0]); ln++ {
			nc += 1
			line = append(line,lines[0][ln])
		}
		rec[0] = line
		// add uma linha vazia
		var lnvazio []string
		for z := 0; z < nc; z++ {
			lnvazio = append(lnvazio," ")
		}
		rec[1] = lnvazio
	}
	return rec
}

func IMMTempDir() (string,error) {
	return ioutil.TempDir(os.TempDir(),"imm")
}

func WhichOS() (string) {
	return runtime.GOOS
}

func ELinux() (bool) {
	eLinux := false
	if WhichOS() == "linux" {
		eLinux = true
	}
	return eLinux
}

func GetUserName() (string) {
	user := "USERNAME"
	if ELinux() {
		user = "USER"
	}
	return os.Getenv(user)
}

// mais eficiente
func EscrevaString2Mem(texto *string,linha string) {
	*texto += linha
}

func EscrevaString(fo *os.File,linha string)(bool) {
	_,err := io.WriteString(fo,linha)
	if err != nil { return false }
	return true
}

func Ano() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s",ja[0:4])
}

func Hoje() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s/%s/%s %s",ja[0:4],ja[5:7],ja[8:10],ja[11:23])
}

func GetFullPathCSV(Tabela string) (string) {
	return filepath.Join(os.TempDir(),fmt.Sprintf("%s.csv",Tabela))
}

func Merge(fnamein1 string,fnamein2 string,fnameout string) (string) {
	buf := bytes.NewBuffer(nil)
	// Leia primeira parte
	f,_ := os.Open(fnamein1)
	io.Copy(buf,f)
	f.Close()
	// Leia segunda parte
	f,_ = os.Open(fnamein2)
	io.Copy(buf,f)
	f.Close()
	// Cria novo arquivo
	fo,err := os.Create(fnameout)
	if err != nil {
		fnameout = ""
	} else {
		if !EscrevaString(fo,string(buf.Bytes())) {
			fnameout = ""
		}
		fo.Close()
	}
	return fnameout
}

func Exists(path string)(bool) {
	_,err := os.Stat(path)
	if err == nil { return true }
	if os.IsNotExist(err) { return false }
	return false
}

func ProcessCSV(path string,Asp ...bool) ([][]string, error) {
	var line []string
	var err error
	asp := true
	if len(Asp) > 0 {
		asp = Asp[0]
	}
	if asp {
		line,_,err = readLines(path)
	} else {
		line,err = readLinesSemAspas(path)
	}
	if err != nil {
		fmt.Println(err)
	}
	lines := make([][]string,len(line))
	for l := 0; l < len(line); l++ {
		reader := csv.NewReader(strings.NewReader(line[l]))
		if asp {
			reader.Comma = ','
		} else {
			reader.Comma = ';'
		}
		reader.LazyQuotes = true
		reader.FieldsPerRecord = 0 // mesmo numero de campos
		records,_ := reader.Read()
		var oneLine []string
		for r := 0; r < len(records); r++ {
			campo := records[r]
			if len(campo) > 1 {
				campo = strings.TrimSpace(campo)
			}
			if len(campo) == 0 {
				campo = " "
			}
			oneLine = append(oneLine,campo)
		}
		lines[l] = oneLine
	}
	return lines,err
}

func readLinesSemAspas(path string) ([]string,error) {
	bom := isBOMFile(path)
	file,err := os.Open(path)
	if err != nil {
		return nil,err
	}
	defer file.Close()
	var first bool = true
	var lines []string
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		ss := scanner.Text()
		if first {
			first = false
			sln := ss
			if bom {
				sln = ss[3:]
			}
			lines = append(lines,sln)
		} else {
			lines = append(lines,ss)
		}
	}
	return lines,scanner.Err()
}

func readLines(path string) ([]string,int,error) {
	bom := isBOMFile(path)
	file,err := os.Open(path)
	if err != nil {
		return nil,0,err
	}
	var nj int = 0
	defer file.Close()
	var first bool = true
	var lines []string
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		ss := scanner.Text()
		if first {
			first = false
			sln := ss
			if bom {
				sln = ss[3:]
			}
			lines = append(lines,sln)
		} else {
			if ss[len(ss)-1:] != "\"" {
				if scanner.Scan() {
					ss2 := scanner.Text()
					lines = append(lines,ss+ss2)
					nj++
				} else {
					lines = append(lines,ss)
				}
			} else {
				lines = append(lines,ss)
			}
		}
	}
	return lines,nj,scanner.Err()
}

func isBOMFile(path string) bool {
	ok := false
	file,err := os.Open(path)
	if err != nil {
		return ok
	}
	defer file.Close()
	rd := make([]byte,3)
	file.Read(rd)
	if rd[0] == 0xEF && rd[1] == 0xBB && rd[2] == 0xBF {
		ok = true
	}
	return ok
}

func NextExtension() (string) {
	ext := new([max]string)
	ext[0] = "_"
	ext[1] = "__"
	ext[2] = "$"
	ext[3] = "_$"
	ext[4] = "$_"
	ext[5] = "#"
	ext[6] = "##"
	ext[7] = "#_"
	ext[8] = "_#"
	ext[9] = "$#"
	s2 := rand.NewSource(time.Now().UnixNano())
	r2 := rand.New(s2)
	n := r2.Intn(max)
	ne := n % (max/2)
	vale := ext[ne]
	if vale == antes {
		ne = n / (max/2)
		vale = ext[ne]
	}
	antes = vale
	return vale
}

func ReadConfDB(cfg string) (INFODB) {
	var infoDB INFODB
	file,err := os.Open(cfg)
	if err == nil {
		defer file.Close()
		var rd []byte
		rd,err= ioutil.ReadAll(file)
		if err == nil {
			var rec string
			rec = string(rd[:])
			lines := strings.Split(rec,"\n")
			for l := 0; l < len(lines); l++ {
				if !strings.Contains(lines[l],"#") {
					parts := strings.Split(lines[l],"=")
					if len(parts) == 2 {
						tipo := strings.ToUpper(parts[0])
						if tipo == "DBNAME" {
							infoDB.DBName = parts[1]
							fmt.Printf("Base: %s\n",infoDB.DBName)
						} else if tipo == "USER"{
							infoDB.User = parts[1]
							fmt.Printf("Usuário: %s\n",infoDB.User)
						} else if tipo == "PWD"{
							infoDB.Password = parts[1]
						} else if tipo == "HOST"{
							host,err := net.LookupHost(parts[1])
							if err == nil {
								infoDB.Host = host[0]
								fmt.Printf("Host: %s %s\n",infoDB.Host,parts[1])
							} else {
								infoDB.Err = err
							}
						} else if tipo == "SSLMODE"{
							infoDB.sslmode = parts[1]
							fmt.Printf("SSLMode: %s\n",infoDB.sslmode)
						} else if tipo == "PORTA"{
							infoDB.Port,_ = strconv.Atoi(parts[1])
							fmt.Printf("Porta: %d\n",infoDB.Port)
						}
					}
				}
			}
		}
	}
	return infoDB
}

// Abra um Base PSQL
func ConecPSQL(infoDB INFODB) (*sql.DB,error) {
	conf := fmt.Sprintf("host=%s port=%d dbname=%s user=%s password=%s sslmode=%s",infoDB.Host,infoDB.Port,infoDB.DBName,infoDB.User,infoDB.Password,infoDB.sslmode)
	db,err := sql.Open("postgres",conf)
	return db,err
}

/*
func readConfDB(cfg string) (INFODB) {
	var InfoDB INFODB
	InfoDB.Err = nil
	InfoDB.Host = ""
	InfoDB.DBName = ""
	InfoDB.User = ""
	InfoDB.Password = ""
	InfoDB.sslmode = ""
	InfoDB.Port = 0
	file,err := os.Open(cfg)
	if err == nil {
		defer file.Close()
		var rd []byte
		rd,err= ioutil.ReadAll(file)
		if err == nil {
			var rec string
			rec = string(rd[:])
			lines := strings.Split(rec,"\n")
			for l := 0; l < len(lines); l++ {
				if !strings.Contains(lines[l],"#") {
					parts := strings.Split(lines[l],"=")
					if len(parts) == 2 {
						tipo := strings.ToUpper(parts[0])
						if tipo == "DBNAME" {
							InfoDB.DBName = parts[1]
							fmt.Printf("Base: %s\n",InfoDB.DBName)
						} else if tipo == "USER"{
							InfoDB.User = parts[1]
							fmt.Printf("Usuário: %s\n",InfoDB.User)
						} else if tipo == "PWD"{
							InfoDB.Password = parts[1]
						} else if tipo == "HOST"{
							host,err := net.LookupHost(parts[1])
							if err == nil {
								InfoDB.Host = host[0]
								fmt.Printf("Host: %s %s\n",InfoDB.Host,parts[1])
							}
						} else if tipo == "SSLMODE"{
							InfoDB.sslmode = parts[1]
							fmt.Printf("SSLMode: %s\n",InfoDB.sslmode)
						} else if tipo == "PORTA"{
							InfoDB.Port,_ = strconv.Atoi(parts[1])
							fmt.Printf("Porta: %d\n",InfoDB.Port)
						}
					}
				}
			}
		}
	}
	InfoDB.Err = err
	return InfoDB
}


// Abra um Base PSQL
func conecPSQL(infoDB INFODB) (*sql.DB,error) {
	conf := fmt.Sprintf("Host=%s Port=%d DBName=%s User=%s Password=%s sslmode=%s",infoDB.Host,infoDB.Port,infoDB.DBName,infoDB.User,infoDB.Password,infoDB.sslmode)
	fmt.Println(conf)
	db,err := sql.Open("postgres",conf)
	return db,err
}
*/

func GetFnameOnly(fname string) (string) {
	var fnonly string = ""
	separador := fmt.Sprintf("%c",os.PathSeparator)
	parts := strings.Split(fname,separador)
	if len(parts) > 1 {
		fnonly = parts[len(parts)-1]
	} else {
		fnonly = parts[0]
	}
	return fnonly
}

func Now() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s %s/%s/%s",ja[11:19],ja[8:10],ja[5:7],ja[0:4])
}

func AnoMesDia() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s%s%s",ja[0:4],ja[5:7],ja[8:10])
}

func ReadFileUTF8(path string) (string,error) {
	var rec string
	file,err := os.Open(path)
	if err == nil {
		defer file.Close()
		var rd []byte
		rd,err = ioutil.ReadAll(file)
		if err == nil {
			if rd[0] == 0xEF && rd[1] == 0xBB && rd[2] == 0xBF {
				rec = string(rd[3:])
			} else {
				rec = string(rd[:])
			}
		}
	}
	return rec,err
}

func writeLines(lines [][]string, path string) error {
	file,err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()
	var sep string = ""
	var first bool = true
	var headSkip []bool
	w := bufio.NewWriter(file)
	for nl := 0; nl < len(lines); nl++ {
		var newLine string = ""
		parts := lines[nl]
		if first {
			first = false
			for p := 0; p < len(parts); p++ {
				if strings.Contains(parts[p],"$InternalId") {
					headSkip = append(headSkip,true)
				} else {
					if strings.Contains(parts[p],"$Path") {
						headSkip = append(headSkip,true)
					} else {
						headSkip = append(headSkip,false)
						if len(newLine) == 0 {
							newLine = fmt.Sprintf("%s",parts[p])
						} else {
							newLine = fmt.Sprintf("%s,%s",newLine,parts[p])
						}
					}
				}
			}
			sep = ""
		} else {
			if len(parts) != len(headSkip) {
				newLine = fmt.Sprintf("***Erro Header*** Linha %d - p = %d h = %d\n",nl,len(parts),len(headSkip))
				fmt.Fprintln(w,newLine)
				return errors.New(newLine)
			} else {
				for p := 0; p < len(headSkip); p++ {
					if !headSkip[p] {
						campo := parts[p]
						if len(campo) == 0 {
							campo = " "
						}
						if len(newLine) == 0 {
							newLine = fmt.Sprintf("%s",campo)
						} else {
							newLine = fmt.Sprintf("%s;%s",newLine,campo)
						}
					}
				}
				sep = ";"
			}
		}
		if len(newLine) > 0 {
			newLine = fmt.Sprintf("%s%s",newLine,sep)
			if nl == len(lines)-1 {
				fmt.Fprintf(w,"%s",newLine)
			} else {
				fmt.Fprintln(w,newLine)
			}
		}
	}
	return w.Flush()
}

func renameFile(file1 string,file2 string) error {
	if _,err := os.Stat(file2); err == nil {
		if err = os.Remove(file2); err != nil {
			return err
		}
	}
	return os.Rename(file1,file2)
}

func GetHostName() (error,string) {
	host,err := os.Hostname()
	return err,host
}

func GetHostNameREGER() (error,string) {
	err,host := GetHostName()
	if err == nil {
		if !strings.Contains(host,".") {
			prefixo := strings.ToLower(host[:2])
			if prefixo == "rr"||prefixo == "rj"||prefixo == "rb"||prefixo == "rf" {
				host = fmt.Sprintf("%s.reger.ons",host)
			}
		}
	}
	return err,host
}

func RecrieCSV(arqcsv string,ren ...bool) error {
	mv := true
	if len(ren) > 0 {
		mv = ren[0]
	}
	lines,err := ProcessCSV(arqcsv)
	if err == nil {
		newFile := fmt.Sprintf("%s%s",arqcsv,NextExtension())
		err = writeLines(lines,newFile)
		if err == nil {
			if mv {
				err = renameFile(newFile,arqcsv)
			}
		}
	}
	return err
}

func RecrieCSVFromRec(path string,rec [][]string,ren ...bool) error {
	mv := true
	if len(ren) > 0 {
		mv = ren[0]
	}
	newFile := fmt.Sprintf("%s_",path)
	err := writeLines(rec,newFile)
	if err == nil {
		if mv {
			err = renameFile(newFile,path)
		}
	}
	return err
}

// Define um Estilo para utilizar nas células do Excel
func DefineStyle(cell *xlsx.Cell,Bold bool,FColor string,AHor string) (*xlsx.Style) {
	cstyle := cell.GetStyle()
	cstyle.Font.Italic = false
	cstyle.Font.Bold = Bold
	cstyle.Font.Size = 10
	cstyle.Font.Name = "Verdana"
	cstyle.Border.Top = "thin"
	cstyle.Border.Left = "thin"
	cstyle.Border.Right = "thin"
	cstyle.Border.Bottom = "thin"
	cstyle.Alignment.Horizontal = AHor
	cstyle.Alignment.Vertical = Center
	cstyle.Fill.PatternType = "solid"
	cstyle.Fill.FgColor = FColor
	return cstyle
}

// Converte CSVS para XLSX
// Utilizar golexcel para gerar o CSV
func Csv2XLSX(csv string,fnxlsx string) (string,error) {
	CSVOnly := GetFnameOnly(csv)
	Parts := strings.Split(strings.ToLower(CSVOnly),".")
	if len(Parts) > 1 {
		if !(Parts[1] == "csv" || Parts[1] == "txt") {
			return fnxlsx,errors.New("Extensão não permitida\n")
		}
	}
	fi,err := os.Open(csv)
	if err != nil {
		return fnxlsx,err
	}
	var nra int = 0 // contador de numero de linha por aba
	var abaCorrente string = ""
	var file *xlsx.File = xlsx.NewFile()
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var cstyle *xlsx.Style
	var cstyle2 *xlsx.Style
	var cstyle3 *xlsx.Style
	var cstyle4 *xlsx.Style
	var mergeGposD int = -1
	var mergeGposL int = -1
	var mergeGcolD int = 5
	var mergeGcolL int = 7
	var mudeCOL0 int = -1
	var mudeCOL1 int = -1
	var posIDOLD int = -1
	var posLIMSUP int = -1
	var posLIMINF int = -1
	var sz float64 = 32.0
	var szLim float64 = 32.0
	defer fi.Close()
	r := bufio.NewReader(fi)
	for {
		var s,_,e  = r.ReadLine()
		if e == nil {
			ss := string(s)
			if len(ss) < 2 {
				continue
			}
			if ss[:2] == ";;" {
				continue
			}
			parts := strings.Split(ss,";")
			if len(parts) > 0 {
				if parts[0][0] == '>' {
					posIDOLD = -1
					if parts[0][1:] == "Fim" {
						break
					}
					nra = 0
					abaCorrente = parts[0][1:]
					if abaCorrente != "Log" {
						sheet,err = file.AddSheet(abaCorrente)
						if err != nil {
							return fnxlsx,err
						}
						fmt.Printf("Criado %s\n",abaCorrente)
					}
				} else {
					if abaCorrente == "Log" {
						nra++
					} else if abaCorrente == "Geral" {
						nra++
						if nra == 1 {
							sz = 15.0
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								cell.Value = parts[p]
								cell.NumFmt = "text"
								if p == 0 {
									cstyle = DefineStyle(cell,true,Silver,Center)
								}
								cell.SetStyle(cstyle)
								if p > 1 {
									if strings.Contains(parts[p],"PLACA")    ||
										strings.Contains(parts[p],"LINHA")    ||
										strings.Contains(parts[p],"SRC_ADDR") ||
										strings.Contains(parts[p],"APTITLE")  ||
										strings.Contains(parts[p],"CNF/MUL")  ||
										strings.Contains(parts[p],"OPMSK") {
										sz = 14.0
									} else if strings.Contains(parts[p],"SISTEMA") {
										sz = 24.0
									} else {
										sz = 9.0
									}
								}
								err = sheet.SetColWidth(p,p,sz)
								if err != nil {
									fmt.Printf(err.Error())
								}
							}
						} else {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								if nra == mergeGposD {
									cell.Merge(mergeGcolD,0)
									cell.Value = parts[p]
								} else if nra == mergeGposL {
									cell.Merge(mergeGcolL,0)
									cell.Value = parts[p]
								} else {
									cell.Value = parts[p]
								}
								if p == 0 {
									cstyle2 = DefineStyle(cell,true,Silver,Left)
									cell.SetStyle(cstyle2)
									if strings.Contains(parts[p],"DESCRI") {
										mergeGposD = nra
									} else if strings.Contains(parts[p],"LISTA") {
										mergeGposL = nra
									} else if strings.Contains(parts[p],"IPs") {
										mergeGposL = nra
										cell.Value = "LISTA IPS"
									}
								} else if p == 1 {
									cstyle3 = DefineStyle(cell,false,White,Left)
									cell.SetStyle(cstyle3)
									if nra == mergeGposD {
										for pp := 2; pp < len(parts)-1; pp++ {
											cell = row.AddCell()
											cell.Value = " "
											cell.SetStyle(cstyle3)
										}
										break
									} else if nra == mergeGposL {
										for pp := 2; pp < len(parts)-1; pp++ {
											cell = row.AddCell()
											cell.Value = " "
											cell.SetStyle(cstyle3)
										}
										break
									}
								} else {
									cstyle4 = DefineStyle(cell,false,White,Right)
									val,err := strconv.ParseFloat(parts[p],64)
									if err == nil {
										cell.SetFloat(val)
									} else {
										cstyle4.Alignment.Horizontal = Left
									}
									cell.SetStyle(cstyle4)
								}
							}
						}
					} else if abaCorrente == "Analogicos" {
						nra++
						if nra < 4 {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								if nra == 3 {
									if p == 3 {
										posLIMSUP = p
									} else if p == 4 {
										posLIMINF = p
									}
								}
								cell = row.AddCell()
								cstyle = DefineStyle(cell,true,Silver,Center)
								if p == 0 && nra == 1 {
									cell.Merge(len(parts)-2,0)
									cell.Value = parts[p]
								} else if p == 0 && nra == 2 {
									cell.Merge(0,1)
									cell.Value = parts[p]
								} else if p == 1 && nra == 2 {
									cell.Merge(1,0)
									cell.Value = parts[p]
								} else if p == 3 && nra == 2 {
									cell.Merge(1,0)
									cell.Value = parts[p]
								} else if ((p >= 5 && p <= len(parts)-2) && nra == 2) {
									cell.Merge(0,1)
									cell.Value = parts[p]
									if strings.Contains(parts[p],"ID ") {
										err = sheet.SetColWidth(p,p,32)
										if err != nil {
											fmt.Printf(err.Error())
										} else {
											if p > 0 {
												posIDOLD = p
											}
										}
									} else if strings.Contains(parts[p],"Agente") {
										err = sheet.SetColWidth(p,p,10)
										if err != nil {
											fmt.Printf(err.Error())
										}
									} else if strings.Contains(parts[p],"Identificador") {
										err = sheet.SetColWidth(p,p,40)
										if err != nil {
											fmt.Printf(err.Error())
										}
									} else if strings.Contains(parts[p],"Grupo") {
										cell.Value = "Grupo OCR"
										err = sheet.SetColWidth(p,p,16)
										if err != nil {
											fmt.Printf(err.Error())
										}
									}
								} else {
									cell.Value = parts[p]
								}
								cell.NumFmt = "text"
								cell.SetStyle(cstyle)
								if nra == 1 {
									if p ==  0 {
										sz = 32.0
									} else {
										sz = 12.0
									}
									err = sheet.SetColWidth(p,p,sz)
									if err != nil {
										fmt.Printf(err.Error())
									}
								}
							}
						} else {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								if p == 0 {
									cell.NumFmt = "text"
									cell.Value = parts[p]
									cstyle2 = DefineStyle(cell,true,White,Left)
									cell.SetStyle(cstyle2)
								} else {
									val,err := strconv.ParseFloat(parts[p],64)
									if err == nil {
										cell.SetFloat(val)
										cstyle3 = DefineStyle(cell,false,White,Right)
										cell.SetStyle(cstyle3)
									} else {
										var align string = Left
										cell.NumFmt = "text"
										if strings.Contains(parts[p],Pipe) {
											align = Right
											if mudeCOL0 == -1 && mudeCOL1 == -1 {
												err = sheet.SetColWidth(p,p,szLim)
												if err != nil {
													fmt.Printf(err.Error())
												} else {
													mudeCOL0 = p
												}
											} else if mudeCOL0 >= 0 && mudeCOL0 != p && mudeCOL1 == -1 {
												err = sheet.SetColWidth(p,p,szLim)
												if err != nil {
													fmt.Printf(err.Error())
												} else {
													mudeCOL1 = p
												}
											}
										}
										xval := parts[p]
										if p == posLIMSUP || p == posLIMINF {
											xp := strings.Split(xval,Pipe)
											lxp := len(xp)
												if lxp > 1 {
												for x := 0; x < lxp; x++ {
													pos := strings.Index(xp[x],".")
													if pos > 0 {
														xp[x] = xp[x][:pos]
													}
												}
												if lxp == 2 {
													xval = fmt.Sprintf("%s|%s",xp[0],xp[1])
												} else if lxp == 3 {
													xval = fmt.Sprintf("%s|%s|%s",xp[0],xp[1],xp[2])
												} else if lxp == 4 {
													xval = fmt.Sprintf("%s|%s|%s|%s",xp[0],xp[1],xp[2],xp[3])
												}
											}
										}
										cell.Value = xval
										c_cor := White
										if p == posIDOLD {
											c_cor = Yellow
										}
										cstyle4 = DefineStyle(cell,false,c_cor,align)
										cell.SetStyle(cstyle4)
									}
								}
							}
						}
					} else if abaCorrente == "Digitais" {
						nra++
						if nra < 4 {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								cstyle = DefineStyle(cell,true,Silver,Center)
								if p == 0 && nra == 1 {
									cell.Merge(len(parts)-2,0)
									cell.Value = parts[p]
								} else if ((p >= 0 && p <= len(parts)-2) && nra == 2) {
									cell.Merge(0,1)
									cell.Value = parts[p]
									if strings.Contains(parts[p],"ID ") {
										err = sheet.SetColWidth(p,p,32)
										if err != nil {
											fmt.Printf(err.Error())
										} else {
											if p > 0 {
												posIDOLD = p
											}
										}
									} else if strings.Contains(parts[p],"Alarm") {
										cell.Value = "Alarme/SOE"
										err = sheet.SetColWidth(p,p,60)
										if err != nil {
											fmt.Printf(err.Error())
										}
									} else if strings.Contains(parts[p],"Identificador") {
										err = sheet.SetColWidth(p,p,40)
										if err != nil {
											fmt.Printf(err.Error())
										}
									} else if strings.Contains(parts[p],"Grupo") {
										cell.Value = "Grupo OCR"
										err = sheet.SetColWidth(p,p,16)
										if err != nil {
											fmt.Printf(err.Error())
										}
									}
								} else {
									cell.Value = parts[p]
								}
								cell.NumFmt = "text"
								cell.SetStyle(cstyle)
								if nra == 1 {
									if p ==  0 {
										sz = 32.0
									} else {
										sz = 12.0
									}
									err = sheet.SetColWidth(p,p,sz)
									if err != nil {
										fmt.Printf(err.Error())
									}
								}
							}
						} else {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								if p == 0 {
									cell.NumFmt = "text"
									cell.Value = parts[p]
									cstyle2 = DefineStyle(cell,true,White,Left)
									cell.SetStyle(cstyle2)
								} else {
									align := Right
									c_cor := White
									if p == posIDOLD {
										c_cor = Yellow
										align = Left
									}
									val,err := strconv.ParseFloat(parts[p],64)
									if err == nil {
										cell.SetFloat(val)
										cstyle3 = DefineStyle(cell,false,c_cor,align)
										cell.SetStyle(cstyle3)
									} else {
										cell.NumFmt = "text"
										cell.Value = parts[p]
										cstyle4 = DefineStyle(cell,false,c_cor,Left)
										cell.SetStyle(cstyle4)
									}
								}
							}
						}
					} else if abaCorrente == "Totalizados" {
						nra++
						if nra < 4 {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								cstyle = DefineStyle(cell,true,Silver,Center)
								if p == 0 && nra == 1 {
									cell.Merge(len(parts)-2,0)
									cell.Value = parts[p]
								} else if p == 0 && nra == 2 {
									cell.Merge(0,1)
									cell.Value = parts[p]
								} else if p == 1 && nra == 2 {
									cell.Merge(1,0)
									cell.Value = parts[p]
								} else if p == 3 && nra == 2 {
									cell.Merge(1,0)
									cell.Value = parts[p]
								} else if ((p >= 5 && p <= len(parts)-2) && nra == 2) {
									cell.Merge(0,1)
									cell.Value = parts[p]
									if strings.Contains(parts[p],"ID ") {
										err = sheet.SetColWidth(p,p,32)
										if err != nil {
											fmt.Printf(err.Error())
										}
									} else if strings.Contains(parts[p],"Fator") {
										err = sheet.SetColWidth(p,p,20)
										if err != nil {
											fmt.Printf(err.Error())
										}
									}
								} else {
									cell.Value = parts[p]
								}
								cell.NumFmt = "text"
								cell.SetStyle(cstyle)
								if nra == 1 {
									if p ==  0 {
										sz = 32.0
									} else {
										sz = 12.0
									}
									err = sheet.SetColWidth(p,p,sz)
									if err != nil {
										fmt.Printf(err.Error())
									}
								}
							}
						} else {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								if p == 0 {
									cell.NumFmt = "text"
									cell.Value = parts[p]
									cstyle2 = DefineStyle(cell,true,White,Left)
									cell.SetStyle(cstyle2)
								} else {
									val,err := strconv.ParseFloat(parts[p],64)
									if err == nil {
										cell.SetFloat(val)
										cstyle3 = DefineStyle(cell,false,White,Right)
										cell.SetStyle(cstyle3)
									} else {
										cell.NumFmt = "text"
										cell.Value = parts[p]
										cstyle4 = DefineStyle(cell,false,White,Left)
										cell.SetStyle(cstyle4)
									}
								}
							}
						}
					} else if abaCorrente == "Controles" {
						nra++
						if nra < 4 {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								cstyle = DefineStyle(cell,true,Silver,Center)
								if p == 0 && nra == 1 {
									cell.Merge(len(parts)-2,0)
									cell.Value = parts[p]
								} else if ((p >= 0 && p <= len(parts)-2) && nra == 2) {
									cell.Merge(0,1)
									cell.Value = parts[p]
									if strings.Contains(parts[p],"ID ") {
										err = sheet.SetColWidth(p,p,32)
										if err != nil {
											fmt.Printf(err.Error())
										}
									}
								} else {
									cell.Value = parts[p]
								}
								cell.NumFmt = "text"
								cell.SetStyle(cstyle)
								if nra == 1 {
									if p ==  0 {
										sz = 32.0
									} else {
										sz = 12.0
									}
									err = sheet.SetColWidth(p,p,sz)
									if err != nil {
										fmt.Printf(err.Error())
									}
								}
							}
						} else {
							row = sheet.AddRow()
							for p := 0; p < len(parts)-1; p++ {
								cell = row.AddCell()
								if p == 0 {
									cell.NumFmt = "text"
									cell.Value = parts[p]
									cstyle2 = DefineStyle(cell,true,White,Left)
									cell.SetStyle(cstyle2)
								} else {
									val,err := strconv.ParseFloat(parts[p],64)
									if err == nil {
										cell.SetFloat(val)
										cstyle3 = DefineStyle(cell,false,White,Right)
										cell.SetStyle(cstyle3)
									} else {
										cell.NumFmt = "text"
										cell.Value = parts[p]
										cstyle4 = DefineStyle(cell,false,White,Left)
										cell.SetStyle(cstyle4)
									}
								}
							}
						}
					}
				}
			}
		} else {
			break
		}
	}
	if len(fnxlsx) > 5 {
		err = file.Save(fnxlsx)
	}
	return fnxlsx,err
}

func LExcel2Memory(pgm string,versao string,xlsFile string,show bool) (int,string) {
	var erro int = -1
	var txtFile string = ""
	var txtSaida string = ""
	var txtSaidaAux string = ""
	fnlower := strings.ToLower(xlsFile)
	pos := strings.Index(fnlower,".xlsx")
	if pos < 0 {
		pos = strings.Index(fnlower,".xlsm")
	}
	if pos >= 0 {
		fOk := Exists(xlsFile)
		if fOk {
			erro = 0
			xlFile,err := xlsx.OpenFile(xlsFile)
			if err != nil {
				erro = 1
			} else {
				var replacer = strings.NewReplacer(";"," ","\n"," ","\t"," ","\r"," ",">"," ","<"," ")
				var abas_posic []int
				var abasmaxcol []int
				var abasmaxrow []int
				var nab int = 0
				var abas []string
				var folha []*xlsx.Sheet
				var linepaba []int
				for s,sheet := range xlFile.Sheets {
					abas = append(abas,sheet.Name)
					linepaba = append(linepaba,0)
					folha = append(folha,sheet)
					abas_posic = append(abas_posic,s)
					abasmaxcol = append(abasmaxcol,sheet.MaxCol)
					abasmaxrow = append(abasmaxrow,sheet.MaxRow)
					nab += 1
				}
				line := fmt.Sprintf(">Log\n%s versao %s\n%s - Processando %s\n",pgm,versao,Hoje(),GetFnameOnly(xlsFile))
				if show {
					fmt.Printf("Processando %s\n",GetFnameOnly(xlsFile))
				}
				EscrevaString2Mem(&txtSaida,line)
				for a := 0; a < nab; a++ {
					f := abas_posic[a]
					if f < 0 {
						continue
					}
					if show {
						fmt.Printf("%2d %-25s %2d %5d\n",f,abas[f],abasmaxcol[f],abasmaxrow[f])
					}
					if abas[f] == "Geral" {
						if abasmaxcol[f] < 2 || abasmaxrow[f] < 2 {
							erro = 15
							break
						}
					}
					line = fmt.Sprintf(">%s\n",abas[f])
					EscrevaString2Mem(&txtSaidaAux,line)
					// processe cada linha
					sheet := folha[f]
					for _,row := range sheet.Rows {
						linepaba[f] += 1
						// processe as colunas desta linha
						for _,cell := range row.Cells {
							valor := cell.String()
							if len(valor) == 0 {
								valor = " "
							} else {
								valor = strings.TrimSpace(replacer.Replace(valor))
								if len(valor) == 0 {
									valor = " "
								}
							}
							valor = fmt.Sprintf("%s;",valor)
							EscrevaString2Mem(&txtSaidaAux,valor)
						}
						EscrevaString2Mem(&txtSaidaAux,"\n")
					}
					line = fmt.Sprintf("%s - Processada %s\n",Hoje(),abas[f])
					EscrevaString2Mem(&txtSaida,line)
				}
				txtaba := "abas"
				if nab < 2 {
					txtaba = "aba"
				}
				line = fmt.Sprintf("Lido %d %s. Linhas por Aba: ",nab,txtaba)
				EscrevaString2Mem(&txtSaida,line)
				for a := 0; a < nab; a++ {
					var fim string = " "
					if a == (nab-1) {
						fim = "\n"
					}
					line = fmt.Sprintf("%s=%d%s",abas[a],linepaba[a],fim)
					EscrevaString2Mem(&txtSaida,line)
				}
				line = fmt.Sprintf("%s - Fim Processamento\n",Hoje())
				EscrevaString2Mem(&txtSaida,line)
				EscrevaString2Mem(&txtSaidaAux,">Fim\n")
			}
		}
	}
	if erro == 0 {
		txtFile = txtSaida + txtSaidaAux;
	}
	return erro,txtFile
}

func LExcel(pgm string,versao string,xlsFile string,txtFile string,show bool) (int,string) {
	var erro int = -1
	var txtSaida string = ""
	var txtSaidaAux string = ""
	fnlower := strings.ToLower(xlsFile)
	pos := strings.Index(fnlower,".xlsx")
	if pos < 0 {
		pos = strings.Index(fnlower,".xlsm")
	}
	if pos >= 0 {
		fOk := Exists(xlsFile)
		if fOk {
			erro = 0
			xlFile,err := xlsx.OpenFile(xlsFile)
			if err != nil {
				erro = 1
			} else {
				if len(txtFile) > 0 {
					txtSaida = fmt.Sprintf("%s_",txtFile)
				} else {
					txtSaida = fmt.Sprintf("%s.txt_",xlsFile[0:pos])
				}
				fo,err := os.Create(txtSaida)
				if err != nil {
					erro = 2
				} else {
					var replacer = strings.NewReplacer(";"," ","\n"," ","\t"," ","\r"," ",">"," ","<"," ")
					var abas_posic []int
					var abasmaxcol []int
					var abasmaxrow []int
					var nab int = 0
					var abas []string
					var folha []*xlsx.Sheet
					var linepaba []int
					for s,sheet := range xlFile.Sheets {
						abas = append(abas,sheet.Name)
						linepaba = append(linepaba,0)
						folha = append(folha,sheet)
						abas_posic = append(abas_posic,s)
						abasmaxcol = append(abasmaxcol,sheet.MaxCol)
						abasmaxrow = append(abasmaxrow,sheet.MaxRow)
						nab += 1
					}
					line := fmt.Sprintf(">Log\n%s versao %s\n%s - Processando %s\n",pgm,versao,Hoje(),GetFnameOnly(xlsFile))
					if !EscrevaString(fo,line) {
						erro = 4
					}
					txtSaidaAux = fmt.Sprintf("%s_",txtSaida)
					foa,err := os.Create(txtSaidaAux)
					if err != nil {
						erro = 5
					}
					for a := 0; a < nab; a++ {
						f := abas_posic[a]
						if f < 0 {
							continue
						}
						if show {
							fmt.Printf("%2d %-25s %2d %5d\n",f,abas[f],abasmaxcol[f],abasmaxrow[f])
						}
						if abas[f] == "Geral" {
							if abasmaxcol[f] < 2 || abasmaxrow[f] < 2 {
								erro = 15
								break
							}
						}
						line = fmt.Sprintf(">%s\n",abas[f])
						if !EscrevaString(foa,line) {
							erro = 6
							break
						}
						// processe cada linha
						sheet := folha[f]
						for _,row := range sheet.Rows {
							if erro != 0 { break }
							linepaba[f] += 1
							// processe as colunas desta linha
							for _,cell := range row.Cells {
								valor := cell.String()
								if len(valor) == 0 {
									valor = " "
								} else {
									valor = strings.TrimSpace(replacer.Replace(valor))
									if len(valor) == 0 {
										valor = " "
									}
								}
								valor = fmt.Sprintf("%s;",valor)
								if !EscrevaString(foa,valor) {
									erro = 7
									break
								}
							}
							if !EscrevaString(foa,"\n") {
								erro = 8
								break
							}
						}
						line = fmt.Sprintf("%s - Processada %s\n",Hoje(),abas[f])
						if !EscrevaString(fo,line) {
							erro = 9
							break
						}
					}
					txtaba := "abas"
					if nab < 2 {
						txtaba = "aba"
					}
					line = fmt.Sprintf("Lido %d %s. Linhas por Aba: ",nab,txtaba)
					if !EscrevaString(fo,line) {
						erro = 10
					}
					for a := 0; a < nab; a++ {
						var fim string = " "
						if a == (nab-1) {
							fim = "\n"
						}
						line = fmt.Sprintf("%s=%d%s",abas[a],linepaba[a],fim)
						if !EscrevaString(fo,line) {
							erro = 11
						}
					}
					line = fmt.Sprintf("%s - Fim Processamento\n",Hoje())
					if !EscrevaString(fo,line) {
						erro = 12
					}
					if !EscrevaString(foa,">Fim\n") {
						erro = 13
					}
					foa.Close()
				}
				fo.Close()
			}
		}
	}
	if erro == 0 {
		txtfile := Merge(txtSaida,txtSaidaAux,txtFile)
		if len(txtfile) == 0 {
			erro = 14
		} else {
			txtFile = txtfile
		}
	}
	if os.Remove(txtSaida) != nil {
	}
	if os.Remove(txtSaidaAux) != nil {
	}
	return erro,txtFile
}

