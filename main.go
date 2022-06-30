package main

import (
	"encoding/csv"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/olekukonko/tablewriter"
	"github.com/urfave/cli/v2"
	"github.com/xuri/excelize/v2"
)

func main() {
	app := &cli.App{
		Name:  "go-xls2nb",
		Usage: "",
		Flags: []cli.Flag{
			&cli.BoolFlag{
				Name:    "debug",
				Usage:   "Debug output",
				Value:   false,
				Aliases: []string{"d"},
			},
		},
		Commands: []*cli.Command{
			{
				Name:    "sheets",
				Usage:   "Output list sheets in XLS files",
				Aliases: []string{"l"},
				Action:  lists_sheets,
				Flags: []cli.Flag{
					&cli.StringFlag{Name: "file", Required: true, Aliases: []string{"f"}},
				},
			},
			{
				Name:    "rows",
				Usage:   "Output rows in sheet",
				Aliases: []string{"r"},
				Action:  rows_sheets,
				Flags: []cli.Flag{
					&cli.StringFlag{Name: "file", Required: true, Aliases: []string{"f"}},
					&cli.StringFlag{Name: "sheet", Required: true, Aliases: []string{"s"}},
					&cli.StringFlag{Name: "header", Usage: "Добавить заголовок в формете (1,2,3,...) (необязательно)", Value: "", Aliases: []string{"t"}},
					&cli.StringFlag{Name: "output", Usage: "Формат вывода", Value: "tbl", Aliases: []string{"o"}},
				},
			},
		},
	}

	err := app.Run(os.Args)
	if err != nil {
		log.Fatal(err)
	}
}

func lists_sheets(context *cli.Context) error {
	f, err := excelize.OpenFile(context.String("file"))
	if err != nil {
		return err
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()
	table := tablewriter.NewWriter(os.Stdout)
	table.SetHeader([]string{"Index", "List name"})
	sheetsList := f.GetSheetList()
	for i, list := range sheetsList {
		table.Append([]string{strconv.Itoa(i), list})
	}

	table.Render()

	return nil
}

func normalize_row(row *[]string, count_cols int) {
	for i := len(*row); i < count_cols; i++ {
		*row = append(*row, "")
	}
}

func rows_sheets(context *cli.Context) error {
	f, err := excelize.OpenFile(context.String("file"))
	if err != nil {
		return err
	}
	defer func() {
		// Закрыть таблицу.
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	// Получить все столбцы в листе
	cols, err := f.GetCols(context.String("sheet"))
	if err != nil {
		return err
	}
	max_count_cols := len(cols)

	// Получить все строки в листе
	rows, err := f.GetRows(context.String("sheet"))
	if err != nil {
		return err
	}

	// Привести все строки к одному размеру
	var normolized_rows [][]string
	for _, row := range rows {
		normalize_row(&row, max_count_cols)
		normolized_rows = append(normolized_rows, row)
	}

	var header []string
	if context.String("header") != "" {
		header = strings.Split(context.String("header"), ",")
	}

	if context.String("output") == "csv" {
		wr := csv.NewWriter(os.Stdout)

		if context.String("header") != "" {
			wr.Write(header)
		}

		for _, row := range normolized_rows {
			wr.Write(row)
		}
		wr.Flush()
	} else {
		table := tablewriter.NewWriter(os.Stdout)
		table.SetAlignment(tablewriter.ALIGN_LEFT)
		table.SetAutoMergeCells(true)
		//table.SetRowLine(true)
		if context.String("header") != "" {
			table.SetHeader(header)
		}
		table.AppendBulk(normolized_rows)
		table.Render()
	}

	return nil
}
