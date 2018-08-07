/*
** Arquivo: util.go by Cananéa
** Atualizado: 18 de Julho de 2018
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
	"time"
	"bytes"
	"github.com/tealeg/xlsx"
)

// Constantes
const Silver = "C0C0C0"
const White  = "FFFFFF"
const Yellow = "FFFFCC"
const Right  = "right"
const Center = "center"
const Left	 = "left"
const Pipe	 = "|"

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
										cell.Value = parts[p]
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
									val,err := strconv.ParseFloat(parts[p],64)
									if err == nil {
										cell.SetFloat(val)
										cstyle3 = DefineStyle(cell,false,White,Right)
										cell.SetStyle(cstyle3)
									} else {
										cell.NumFmt = "text"
										cell.Value = parts[p]
										c_cor := White
										if p == posIDOLD {
											c_cor = Yellow
										}
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
								} else if ((p >= 0 && p <= len(parts)-2) && nra == 2) {
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

func TimeZero()(time.Time) {
	return time.Unix(0,0)
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

func AddString2Mem(texto *string,linha string) {
	var buffer bytes.Buffer
	buffer.WriteString(*texto)
	buffer.WriteString(linha)
	*texto = buffer.String()
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

func GetFnameNoExtension(fname string) (string) {
	parts := strings.Split(GetFnameOnly(fname),".")
	return parts[0]
}

func Hoje() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s/%s/%s %s",ja[0:4],ja[5:7],ja[8:10],ja[11:23])
}

func Ano() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s",ja[0:4])
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

func Now() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s %s/%s/%s",ja[11:19],ja[8:10],ja[5:7],ja[0:4])
}

func Agora() (string) {
	now := time.Now()
	ja := fmt.Sprintf("%s",now)
	return fmt.Sprintf("%s",ja[11:16])
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
				const maxaba = 20
				abas_posic := [maxaba]int{-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1}
				abasmaxcol := [maxaba]int{-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1}
				abasmaxrow := [maxaba]int{-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1}
				var nab int = 0
				var abas [maxaba] string
				var folha [maxaba] *xlsx.Sheet
				var linepaba [maxaba] int
				for s,sheet := range xlFile.Sheets {
					if nab < maxaba - 1 {
						abas[nab] = sheet.Name
						folha[nab] = sheet
						abas_posic[nab] = s
						abasmaxcol[nab] = sheet.MaxCol
						abasmaxrow[nab] = sheet.MaxRow
						nab += 1
					} else {
						erro = 3
					}
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
						fmt.Printf("%2d %-15s %2d %5d\n",f,abas[f],abasmaxcol[f],abasmaxrow[f])
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
					const maxaba = 20
					abas_posic := [maxaba]int{-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1}
					abasmaxcol := [maxaba]int{-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1}
					abasmaxrow := [maxaba]int{-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1}
					var nab int = 0
					var abas [maxaba] string
					var folha [maxaba] *xlsx.Sheet
					var linepaba [maxaba] int
					for s,sheet := range xlFile.Sheets {
						if nab < maxaba - 1 {
							abas[nab] = sheet.Name
							folha[nab] = sheet
							abas_posic[nab] = s
							abasmaxcol[nab] = sheet.MaxCol
							abasmaxrow[nab] = sheet.MaxRow
							nab += 1
						} else {
							erro = 3
						}
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
							fmt.Printf("%2d %-15s %2d %5d\n",f,abas[f],abasmaxcol[f],abasmaxrow[f])
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

func IMMTempDir() (string,error) {
	return ioutil.TempDir(os.TempDir(),"imm")
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

func GetFullPathCSV(Tabela string) (string) {
	return filepath.Join(os.TempDir(),fmt.Sprintf("%s.csv",Tabela))
}

func CrieLOG(logname string) (*os.File,error,string) {
	namelog := filepath.Join(os.TempDir(),fmt.Sprintf("%s_%s",GetUserName(),logname))
	file,err := os.Create(namelog)
	return file,err,namelog
}

func EscreveLOG(w *bufio.Writer,msg string) {
	fmt.Fprintf(w,"%s\n",msg)
	w.Flush()
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

func ReadLines(path string) ([]string,int,error) {
	return readLines(path)
}

func ReadLinesSemAspas(path string) ([]string,error) {
	return readLinesSemAspas(path)
}

func HeaderOnly(csvin string) ([]string) {
	// csvin - arquivo em formato csv gerado pelo IMM
	var line []string
	lines,_,err := readLines(csvin)
	if err == nil {
		first := true
		for ln := 0; ln < len(lines); ln++ {
			if first {
				line = append(line,lines[ln])
				// add uma linha vazia
				lnvazio := " "
				parts := strings.Split(lines[ln],",")
				for z := 0; z < len(parts)-1; z++ {
					lnvazio = lnvazio + "; "
				}
				lnvazio = lnvazio + ";"
				line = append(line,lnvazio)
				break
			}
		}
	}
	return line
}

func FiltreMUL(csvin string,mul string,inicio bool)([]string) {
	// csvin  - arquivo em formato csv gerado pelo IMM
	// mul    - nome da ligação ICCP ou I61850
	// inicio - true, se está no início da linha
	var rec []string
	lines,err := readLinesSemAspas(csvin)
	if err == nil {
		ncampos := 0
		first := true
		for ln := 0; ln < len(lines); ln++ {
			lineout := ""
			if first {
				first = false
				lineout = lines[ln]
				parts := strings.Split(lines[ln],",")
				ncampos = len(parts)
			} else {
				ok := false
				if inicio {
					if CompareSubString(lines[ln],mul,len(mul)) == 0 { // igual?
						ok = true
					}
				} else {
					if strings.Contains(lines[ln],mul) {
						ok = true
					}
				}
				if ok {
					lineout = lines[ln]
				}
			}
			if len(lineout) > 0 {
				rec = append(rec,lineout)
			}
		}
		if len(rec) == 1 {
			lnvazio := " "
			for z := 0; z < ncampos-1; z++ {
				lnvazio = lnvazio + "; "
			}
			lnvazio = lnvazio + ";"
			rec = append(rec,lnvazio)
		}
	}
	return rec
}

func ProcessCSV(path string) ([]string, error) {
	conteudo,_,err := readLines(path)
	if err != nil {
		return nil,err
	}
	var lines []string
	var sep string = "|"
	for ln := 0; ln < len(conteudo); ln++ {
		r := csv.NewReader(strings.NewReader(conteudo[ln]))
		r.Comma = ','         // delimitador 
		r.FieldsPerRecord = 0 // mesmo numero de campos
		for {
			rec,err := r.Read()
			if err == io.EOF {
				break
			}
			if len(rec) == 0 {
				continue
			}
			oneLine := ""
			for rr := range(rec) {
				campo := rec[rr]
				if len(campo) == 0 {
					campo = " "
				} else {
					xc := strings.Split(campo,"\n")
					if len(xc) > 1 {
						campo = xc[0]
					} else {
						xc := strings.Split(campo,"\r")
						if len(xc) > 1 {
							campo = xc[0]
						}
					}
				}
				if len(campo) > 1 {
					campo = strings.TrimSpace(campo)
				}
				if rr < (len(rec) - 1) {
					oneLine = oneLine + campo + sep
				} else {
					oneLine = oneLine + campo
				}
			}
			lines = append(lines,oneLine)
		}
	}
	return lines,err
}

func writeLines(lines []string, path string) error {
	file,err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()
	var first bool = true
	var headSkip []bool
	w := bufio.NewWriter(file)
	for nl := 0; nl < len(lines); nl++ {
		var newLine string = ""
		parts := strings.Split(lines[nl],"|")
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
			}
		}
		if len(newLine) > 0 {
			fmt.Fprintln(w,newLine)
		}
	}
	return w.Flush()
}

func WriteLines(lines []string, path string) error {
	return writeLines(lines,path)
}

func renameFile(file1 string,file2 string) error {
	if _,err := os.Stat(file2); err == nil {
		if err = os.Remove(file2); err != nil {
			return err
		}
	}
	return os.Rename(file1,file2)
}

func RecrieCSV(arqcsv string,ren ...bool) error {
	mv := true
	if len(ren) > 0 {
		mv = ren[0]
	}
	lines,err := ProcessCSV(arqcsv)
	if err == nil {
		newFile := fmt.Sprintf("%s_",arqcsv)
		err = writeLines(lines,newFile)
		if err == nil {
			if mv {
				err = renameFile(newFile,arqcsv)
			}
		}
	}
	return err
}

func RecrieCSVFromRec(path string,rec []string,ren ...bool) error {
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

