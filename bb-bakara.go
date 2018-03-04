package main

import (
	"os"
	"fmt"
	"log"
	"strings"
	"os/exec"
	_ "github.com/lib/pq"
	"time"
	"flag"
	"github.com/jmoiron/sqlx"
	"github.com/tealeg/xlsx"
	"strconv"
)

type persons_type map[string]int

type dbStruct map[string][]string

var tables = dbStruct{
	//"acc_door": []string{"id", "device_id", "door_no", "door_name"},
	//"acc_levelset": []string{"id", "level_name"},
	//"acc_levelset_door_group": []string{"id", "acclevelset_id", "accdoor_id", "accdoor_no_exp", "accdoor_device_id"},
	//"acc_levelset_emp": []string{"id", "acclevelset_id", "employee_id"},
	//"DEPARTMENTS": []string{"id", "DEPTNAME", "SUPDEPTID"},
	//"Machines": []string{"id", "MachineAlias", "IP", "SerialPort", "Port", "Baudrate", "usercount", "FirmwareVersion",
	//	"sn", "device_name", "subnet_mask", "gateway", "area_id", "acpanel_type"},
	//"personnel_area": []string{"id", "areaid", "areaname", "parent_id"},
	//"personnel_issuecard": []string{"id", "create_time", "UserID_id", "cardno", },
	"acc_monitor_log": []string{"id", "time", "pin", "card_no", "device_id", "device_sn", "device_name", "event_point_name"},
	"USERINFO":        []string{"userid", "USERID", "Badgenumber", "name", "Gender", "BIRTHDAY", "CardNo", "lastname", "identitycard", "bankcode1",},
}

type Config struct {
	dbAdapter      string
	connString     string
	outputPath     string
	mdbPath        string
	doNotCalculate bool
	doNotLoadMdb   bool
	onlyLoadMdb    bool
	makor          string
	noMakor        bool
	month          int
	year           int
}

func main() {
	cfg := getConfig()

	db := openDb(cfg.dbAdapter, cfg.connString)
	defer db.Close()

	// Prices *HAVE* to be loaded *BEFORE* db
	if !cfg.noMakor {
		loadPrices(db, cfg)
	}

	if !cfg.doNotLoadMdb {
		loadDB(cfg.mdbPath, tables, db, cfg)
		if cfg.onlyLoadMdb {
			os.Exit(0)
		}
	}

	if !cfg.doNotCalculate {
		calculate(cfg.mdbPath, tables, db, cfg)
		calculateMoney(db, cfg)
		calculateTotalsPerMeal(db, cfg)
		statistics(db, cfg)
	}
}

type controller struct {
	chipName   string
	controller string
}

var controllerMap = map[string]*controller{
	"חדר אוכל": {
		controller: "ארועים",
		chipName:   "חדר אוכל קטן",},
	"רחבה סעודות": {
		controller: "ארועים",
		chipName:   "רחבת סעודות 1",},
	"רמפה": {
		controller: "ארועים",
		chipName:   "ק.0 מדרגות גימס סעודות",},
	"סעודות עולם אירועים מערב": {
		controller: "חדר אמנון פעיל קומה ",
		chipName:   "ק.0 ארועים מערב סעודות",},
	"ק.3 .מזרכ לובי -מסדרון. מסדרון רב": {
		controller: "מסדרון רב",
		chipName:   "ק.3 מזרח לובי - מסדרון",},
	"סוכה 1": {
		controller: "מסדרון רב",
		chipName:   "סוכה 1",},
	"סוכה 2": {
		controller: "מדרגות מזרח 3 נוכחות",
		chipName:   "סוכה 2",},
	"סוכה 3": {
		controller: "מסדרון רב",
		chipName:   "סוכה 3",},
	"סוכה 4": {
		controller: "מדרגות  קומה 3 מערב ",
		chipName:   "סוכה 4",},
	"סוכה 5": {
		controller: "מדרגות  קומה 3 מערב ",
		chipName:   "סוכה 5",},
	"סוכה 6": {
		controller: "מדרגות  קומה 3 מערב ",
		chipName:   "סוכה 6",},
}

var daysMap = map[string]int{
	"א'": 0,
	"ב'": 1,
	"ג'": 2,
	"ד'": 3,
	"ה'": 4,
	"ו'": 5,
	"ש'": 6,
}

func execSql(db *sqlx.DB, sql string, message string) {
	_, err := db.Exec(sql)
	if err != nil {
		log.Fatal(message, err)
		os.Exit(-1)
	}

}

func loadPrices(db *sqlx.DB, cfg *Config) {
	const (
		// 0 - ignored
		excel_date      = 1
		excel_day_name  = 2
		excel_double    = 3
		excel_youth     = 4
		excel_meal_type = 5
		excel_meal_name = 6
		// 7 - ignored
		excel_hours      = 8
		excel_controller = 9
		excel_price      = 10
		// 11 - ignored
		excel_veg = 12
		// 13 - ignored
		excel_kli = 14
	)
	printCommand("loadPrices", cfg.makor)
	xlFile, err := xlsx.OpenFile(cfg.makor)
	if err != nil {
		log.Fatal(err)
		os.Exit(-1)
	}

	db.Exec("DROP TABLE IF EXISTS prices;")
	db.Exec(`
		CREATE TABLE prices (
			day INTEGER,
			dow INTEGER,
			p2 BOOLEAN,
			youth BOOLEAN,
			income TEXT,
			meal TEXT,
			start TEXT,
			finish TEXT,
			controller TEXT,
			chip_name TEXT,
			price INTEGER,
			vegetarian INTEGER,
			kli INTEGER
		)`)
	sheet := xlFile.Sheets[0]
	for _, row := range sheet.Rows {
		cells := row.Cells
		//fmt.Println(cells)
		if cells[1].Value == "" {
			continue
		}
		if _, err = strconv.Atoi(cells[1].Value); err != nil {
			continue
		}

		day, _ := strconv.Atoi(cells[excel_date].Value)
		dow, ok := daysMap[cells[excel_day_name].Value]
		if !ok {
			fmt.Errorf("### Unknown DOW: %s\n", cells[2].Value)
			os.Exit(-1)
		}
		p2v := cells[excel_double].Value
		if p2v != "כן" && p2v != "לא" {
			fmt.Errorf("### Unknown p2: %s\n", p2v)
			os.Exit(-1)
		}
		p2 := p2v == "כן"
		yv := cells[excel_youth].Value
		if yv != "כן" && yv != "לא" {
			fmt.Errorf("### Unknown youth: %s\n", yv)
			os.Exit(-1)
		}
		youth := yv == "כן"
		income := strings.TrimSpace(strings.Replace(cells[excel_meal_type].Value, "'", "׳", -1))
		label := strings.TrimSpace(strings.Replace(cells[excel_meal_name].Value, "'", "׳", -1))
		if income == "" {
			income = label
		}
		s := strings.Split(cells[excel_hours].Value, "-")
		start, end := strings.Replace(s[0], ".", ":", -1), strings.Replace(s[1], ".", ":", -1)
		priceChip, _ := strconv.Atoi(cells[excel_price].Value)
		priceVegetarian, _ := strconv.Atoi(cells[excel_veg].Value)
		priceKli, _ := strconv.Atoi(cells[excel_kli].Value)
		for _, chipName := range strings.Split(cells[excel_controller].Value, ";") {
			if chipName == "" {
				continue
			}
			controller, ok := controllerMap[strings.TrimSpace(chipName)]
			if !ok {
				fmt.Errorf("### Unable to find chip: %s\n", chipName)
				os.Exit(-1)
			}
			query := fmt.Sprintf(`
				INSERT INTO prices (day, dow, p2, youth, income, meal, start, finish, controller, chip_name, price, vegetarian, kli)
				VALUES(%d, %d, %t, %t, '%s', '%s', '%s', '%s', '%s', '%s', %d, %d, %d);
			`, day, dow, p2, youth, income, label, start, end, controller.controller, controller.chipName, priceChip, priceVegetarian, priceKli)
			if _, err = db.Exec(query); err != nil {
				log.Fatal("Query: ", query, "\n\nError: ", err)
				os.Exit(-1)
			}
		}
	}
}

func getConfig() *Config {

	var doNotLoadMdb = flag.Bool("x", false, "Do not reload MDB file, use the existing data")
	var onlyLoadMdb = flag.Bool("X", false, "ONLY reload MDB file")
	var mdbPath = flag.String("m", "ZKAccess.mdb", "Path to an MDB file")
	var outputPath = flag.String("o", "", "** mandatory ** Output file path")
	var month = flag.Int("d", 0, "Month to create report for")
	var year = flag.Int("y", 0, "Year to create report for")
	var pUser = flag.String("u", "postgres", "Postgres User")
	var pPassword = flag.String("p", "postgres", "Postgres Password")
	var pHost = flag.String("h", "localhost", "Postgres host")
	var makor = flag.String("i", "", "File with prices")
	var noMakor = flag.Bool("I", false, "Do not load file with prices")
	var doNotCalculate = flag.Bool("C", false, "Do not perform calculations")

	flag.Parse()

	if *outputPath == "" && *onlyLoadMdb == false {
		flag.PrintDefaults()
		log.Fatal("Please supply all mandatory parameters")
	}

	if *month <= 0 || *month > 12 {
		log.Fatal("Bad month number", *month)
	}

	if *pPassword == "" {
		*pPassword = ""
	} else {
		*pPassword = ":" + *pPassword
	}

	return &Config{
		dbAdapter:      "postgres",
		connString:     fmt.Sprintf("postgres://%s%s@%s/zkaccess", *pUser, *pPassword, *pHost),
		outputPath:     *outputPath,
		mdbPath:        *mdbPath,
		doNotLoadMdb:   *doNotLoadMdb,
		onlyLoadMdb:    *onlyLoadMdb,
		makor:          *makor,
		noMakor:        *noMakor,
		doNotCalculate: *doNotCalculate,
		month:          *month,
		year:           *year,
	}
}

func loadDB(dbpath string, tables dbStruct, db *sqlx.DB, cfg *Config) {
	printCommand("LoadDB", "")

	for table := range tables {
		dbTable := strings.ToLower(table)

		db.Exec("ALTER SEQUENCE " + dbTable + "_" + tables[table][0] + "_seq RESTART WITH 1;")

		runCommand("Create table "+table,
			"mdb-schema --drop-table -T "+table+" "+dbpath+" postgres | fgrep -v 'ADD CONSTRAINT' | tr '[:upper:]' '[:lower:]' | psql -U postgres -d zkaccess")
		runCommand("Import table "+table,
			"mdb-export -H -q \\\" -D '%Y-%m-%d %H:%M:%S' "+dbpath+" "+table+"| psql -U postgres -d zkaccess -c 'COPY "+dbTable+" FROM STDIN CSV'")
	}

	// keep only records that belong to our time range
	query := fmt.Sprintf(`
		DELETE FROM acc_monitor_log WHERE time < TIMESTAMP '%d-%02d-01 00:00:00' OR time >= (TIMESTAMP '%d-%02d-01 00:00:00' + interval '1 month');
	`, cfg.year, cfg.month, cfg.year, cfg.month)
	execSql(db, query, "Shorten acc_monitor_log")

	// keep only records that belong to our events
	query = `
		DROP TABLE IF EXISTS events;
		SELECT DISTINCT id, pin, card_no, device_name, event_point_name,
			extract(dow from time)::integer AS dow,
			extract(day from time)::integer AS day,
			to_char(time, 'HH24:MI') AS hm
		INTO events
		FROM acc_monitor_log
		WHERE false
	`
	for _, v := range controllerMap {
		query += fmt.Sprintf(` OR (device_name = '%s' AND event_point_name = '%s')`, v.controller, v.chipName)
	}
	execSql(db, query, "Events")
}

func calculate(dbpath string, tables dbStruct, db *sqlx.DB, cfg *Config) {

	query := `
		DROP TABLE IF EXISTS all_records;
		SELECT DISTINCT e.*, u.name, u.lastname, u.ophone, u.fphone, u.pager, u.city, p.meal, p.price, p.vegetarian, p.kli, p.income, p.p2, p.youth
		INTO all_records
		FROM events e
		LEFT OUTER JOIN userinfo u ON u.badgenumber = e.pin OR u.cardno = e.card_no
		LEFT OUTER JOIN prices p ON e.dow = p.dow AND e.day = p.day AND e.hm BETWEEN p.start AND p.finish
		AND e.device_name = p.controller AND e.event_point_name = p.chip_name
		WHERE (coalesce(u.name, '') || ' ' || coalesce(u.lastname, '')) != ' ';
		-- ignore meals that youth should not be charged
		DELETE FROM all_records a WHERE city = '1' AND NOT youth;
		-- ignore meals that kli should not be charged
		DELETE FROM all_records a WHERE city = '2' AND kli IS NULL;
		DELETE FROM all_records a WHERE city = '2' AND kli = 0;
		-- ignore oraat keva
		DELETE FROM all_records a WHERE a.dow = 5 AND a.fphone = '2' AND a.hm > '15:00';
		DELETE FROM all_records a WHERE a.dow = 6 AND a.fphone = '2';
		`
	execSql(db, query, "All records")

	query = `
		DROP TABLE IF EXISTS meals;
		CREATE TABLE meals AS SELECT * FROM all_records WHERE p2; -- Take all rows with double payments
		INSERT INTO meals
		 	SELECT DISTINCT ON (day, meal, name || ' ' || lastname) * FROM all_records WHERE NOT p2; -- Add rows _WITHOUT_ double payments
		DROP TABLE IF EXISTS final_results;
	`
	execSql(db, query, "Meals")

	query = fmt.Sprintf(`
		SELECT  name "שם",
			lastname "שם משפחה",
			''::text "טלפון",
			ophone "מס כרטיס בהנה""ח",
			income "הכנסה",
			meal "סוג סעודה",
			''::text "פסצקה",
			CASE
				WHEN pager='1' THEN vegetarian 
				WHEN city='2' THEN kli
				ELSE price 
			END "מחיר",
			('%d-%02d-' || day)::DATE "תאריך",
			CASE WHEN pager='1' THEN 'כן' ELSE '' END "צמחוני",
			pin || ',' || card_no "מס' כרטיס/צ'יפ",
			CASE
				WHEN device_name = 'ארועים' THEN
					CASE
						WHEN event_point_name = 'חדר אוכל קטן' THEN 'חדר אוכל'
						WHEN event_point_name = 'רחבת סעודות 1' THEN 'רחבה סעודות'
						WHEN event_point_name = 'ק.0 מדרגות גימס סעודות' THEN 'רמפה'
					END
				WHEN device_name = 'חדר אמנון פעיל קומה ' THEN
					CASE WHEN event_point_name = 'ק.0 ארועים מערב סעודות' THEN 'סעודות עולם אירועים מערב'
					END
				WHEN device_name = 'מסדרון רב' THEN
					CASE
						WHEN event_point_name = 'ק.3 מזרח לובי - מסדרון' THEN 'ק.3 .מזרכ לובי -מסדרון. מסדרון רב'
						WHEN event_point_name = 'סוכה 1' THEN 'סוכה 1'
						WHEN event_point_name = 'סוכה 3' THEN 'סוכה 3'
					END
				WHEN device_name = 'מדרגות מזרח 3 נוכחות' THEN
					CASE
						WHEN event_point_name = 'סוכה 2' THEN 'סוכה 2'
					END
				WHEN device_name = 'מדרגות  קומה 3 מערב' THEN
				CASE
					WHEN event_point_name = 'סוכה 4' THEN 'סוכה 4'
					WHEN event_point_name = 'סוכה 5' THEN 'סוכה 5'
					WHEN event_point_name = 'סוכה 6' THEN 'סוכה 6'
				END
			END "קולט",
			hm
		INTO final_results
		FROM meals
		WHERE meal IS NOT NULL;
		`, cfg.year, cfg.month, )
	execSql(db, query, "Final results")

	for_kolia, _ := db.Queryx(`
		SELECT day, hm, card_no, device_name, event_point_name
		FROM events e
		LEFT OUTER JOIN userinfo u ON u.badgenumber = e.pin OR u.cardno = e.card_no
		WHERE (coalesce(u.name, '') || ' ' || coalesce(u.lastname, '')) = ' '
	`)
	file := xlsx.NewFile()
	defer file.Save(cfg.outputPath + "/for_kolia" + strconv.Itoa(cfg.month) + "-" + strconv.Itoa(cfg.year) + ".xlsx")
	writeSheet(for_kolia, file)
}

func runCommand(description string, command string) {
	printCommand(description, command)
	cmd := exec.Command("sh", "-c", command)
	if out, err := cmd.CombinedOutput(); err != nil {
		printError(err)
		printOutput(out)
		os.Exit(1)
	}
}

func printCommand(meaning string, command string) {
	fmt.Printf("... %s\n", meaning)
	fmt.Printf("==> Executing: %s\n", command)
}

func printError(err error) {
	if err != nil {
		os.Stderr.WriteString(fmt.Sprintf("==> Error: %s\n", err.Error()))
	}
}

func printOutput(outs []byte) {
	if len(outs) > 0 {
		fmt.Printf("==> Output: %s\n", string(outs))
	}
}

func calculateTotalsPerMeal(db *sqlx.DB, cfg *Config) {
	printCommand("calculateTotalsPerMeal", "")
	results, err := db.Queryx(`
		WITH
		run_int AS (
			SELECT meal, day::TEXT AS day, count(1), SUM(CASE WHEN pager = '1' THEN vegetarian ELSE price END)
			FROM meals
			GROUP BY meal, day
			ORDER BY meal, day ASC
		),
		run_tot AS (
			SELECT meal, '== Total =='::TEXT AS day, count(1), SUM(CASE WHEN pager = '1' THEN vegetarian ELSE price END)
			FROM meals
			GROUP BY meal
			ORDER BY meal
		)
		SELECT * FROM run_int
		UNION ALL
		SELECT * FROM run_tot
		ORDER BY meal, day ASC
	`)
	if err != nil {
		log.Fatal(err)
		os.Exit(-1)
	}

	file := xlsx.NewFile()
	defer file.Save(cfg.outputPath + "/meals" + strconv.Itoa(cfg.month) + "-" + strconv.Itoa(cfg.year) + ".xlsx")

	writeSheet(results, file)

}

func calculateMoney(db *sqlx.DB, cfg *Config) {
	printCommand("calculateMoney", "")
	results, err := db.Queryx(`
		SELECT *
		FROM final_results
		ORDER BY "מס כרטיס בהנה""ח", "שם משפחה", "שם", "תאריך";
	`)
	if err != nil {
		log.Fatal(err)
		os.Exit(-1)
	}

	file := xlsx.NewFile()
	defer file.Save(cfg.outputPath + "/report" + strconv.Itoa(cfg.month) + "-" + strconv.Itoa(cfg.year) + ".xlsx")

	writeSheet(results, file)
}

func statistics(db *sqlx.DB, cfg *Config) {
	printCommand("statistics", "")
	// Statistics per input device
	results, err := db.Queryx(`
	WITH a AS (
	SELECT device_name, event_point_name,
	CASE
		WHEN device_name = 'ארועים' THEN
			CASE
				WHEN event_point_name = 'חדר אוכל קטן' THEN 'חדר אוכל'
				WHEN event_point_name = 'רחבת סעודות 1' THEN 'רחבה סעודות'
				WHEN event_point_name = 'ק.0 מדרגות גימס סעודות' THEN 'רמפה'
				ELSE device_name || '--' || event_point_name
			END
		WHEN device_name = 'חדר אמנון פעיל קומה ' THEN
			CASE
				WHEN event_point_name = 'ק.0 ארועים מערב סעודות' THEN 'סעודות עולם אירועים מערב'
				ELSE device_name || '--' || event_point_name
			END
		WHEN device_name = 'מסדרון רב' THEN
			CASE
				WHEN event_point_name = 'ק.3 מזרח לובי - מסדרון' THEN 'ק.3 .מזרכ לובי -מסדרון. מסדרון רב'
				WHEN event_point_name = 'סוכה 1' THEN 'סוכה 1'
				WHEN event_point_name = 'סוכה 3' THEN 'סוכה 3'
				ELSE device_name || '--' || event_point_name
			END
		WHEN device_name = 'מדרגות מזרח 3 נוכחות' THEN
			CASE WHEN event_point_name = 'סוכה 2' THEN 'סוכה 2'
			ELSE device_name || '--' || event_point_name
			END
		WHEN device_name = 'מדרגות  קומה 3 מערב' THEN
			CASE
				WHEN event_point_name = 'סוכה 4' THEN 'סוכה 4'
				WHEN event_point_name = 'סוכה 5' THEN 'סוכה 5'
				WHEN event_point_name = 'סוכה 6' THEN 'סוכה 6'
				ELSE device_name || '--' || event_point_name
			END
		ELSE device_name || '--' || event_point_name
		END "device"
	FROM acc_monitor_log
	WHERE device_name IN ('ארועים', 'חדר אמנון פעיל קומה ', 'מסדרון רב', 'מדרגות  קומה 3 מערב ', 'מדרגות מזרח 3 נוכחות') AND
	      event_point_name IN ('חדר אוכל קטן', 'רחבת סעודות 1', 'ק.0 מדרגות גימס סעודות', 'ק.0 ארועים מערב סעודות', 'ק.3 מזרח לובי - מסדרון',
	      'סוכה 6','סוכה 5','סוכה 4','סוכה 3','סוכה 2','סוכה 1'
	      )
	)
	SELECT count(device), device, device_name, event_point_name
	FROM a
	GROUP BY device, device_name, event_point_name
	`)
	if err != nil {
		log.Fatal(err)
		os.Exit(-1)
	}

	file := xlsx.NewFile()
	defer file.Save(cfg.outputPath + "/chip_reader" + strconv.Itoa(cfg.month) + "-" + strconv.Itoa(cfg.year) + ".xlsx")

	writeSheet(results, file)

	// Price per "מס כרטיס בהנה""ח"
	results, err = db.Queryx(`
	WITH
	run_int AS (
		SELECT ''::TEXT AS "סה""כ", income "הכנסה", ophone "מס כרטיס בהנה""ח", SUM(CASE 
			WHEN pager = '1' THEN vegetarian 
			WHEN city = '2' THEN kli 
			ELSE price 
		END) "מחיר"
		FROM meals
		GROUP BY ophone, income
		ORDER BY "הכנסה"
	),
	run_tot AS (
		SELECT 'סה"כ'::TEXT AS "סה""כ", income "הכנסה", ''::TEXT "מס כרטיס בהנה""ח", SUM(CASE 
			WHEN pager = '1' THEN vegetarian 
			WHEN city = '2' THEN kli 
			ELSE price 
		END) "מחיר"
		FROM meals
		GROUP BY income
		ORDER BY "הכנסה"
	)
	SELECT * FROM run_int
	UNION ALL
	SELECT * FROM run_tot
	ORDER BY "הכנסה", "סה""כ"
	`)
	if err != nil {
		log.Fatal(err)
		os.Exit(-1)
	}

	file1 := xlsx.NewFile()
	defer file1.Save(cfg.outputPath + "/totals" + strconv.Itoa(cfg.month) + "-" + strconv.Itoa(cfg.year) + ".xlsx")

	writeSheet(results, file1)
}

func writeSheet(results *sqlx.Rows, file *xlsx.File) (err error) {
	sheet, err := file.AddSheet("Sheet1") // col bestFit="1"
	if err != nil {
		log.Fatal(err)
		os.Exit(-1)
	}

	row := sheet.AddRow()
	columns, _ := results.Columns()
	for _, h := range columns {
		cell := row.AddCell()
		cell.Value = h
	}
	for results.Next() {
		xRow := sheet.AddRow()
		row, err := results.SliceScan()
		if err != nil {
			log.Fatal(err)
			os.Exit(-1)
		}

		for _, col := range row {
			cell := xRow.AddCell()
			switch col.(type) {
			case float64:
				cell.SetFloat(col.(float64))
			case int64:
				cell.SetInt64(col.(int64))
			case bool:
				cell.SetBool(col.(bool))
			case []byte:
				cell.SetString(string(col.([]byte)))
			case string:
				cell.SetString(col.(string))
			case time.Time:
				cell.SetDate(col.(time.Time))
			case nil:
				cell.SetString(" ")
			default:
				log.Print(col)
			}
		}
	}

	return

}

func openDb(adapter string, conn string) *sqlx.DB {

	db, err := sqlx.Open(adapter, conn)
	if err != nil {
		log.Fatal(err)
		os.Exit(-1)
	}

	return db
}
