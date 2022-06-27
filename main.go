package main

import (
	"fmt"
	"log"
	"os"
	"strconv"

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

	table := tablewriter.NewWriter(os.Stdout)
	table.SetAlignment(tablewriter.ALIGN_LEFT)
	// Получить все строки в листе
	rows, err := f.Rows(context.String("sheet"))
	if err != nil {
		return err
	}
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			fmt.Println(err)
		}
		var rowCells []string
		for _, colCell := range row {
			rowCells = append(rowCells, colCell)
		}
		table.Append(rowCells)
	}
	if err = rows.Close(); err != nil {
		return err
	}

	table.Render()

	return nil
}
