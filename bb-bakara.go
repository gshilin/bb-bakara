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
	"github.com/rach/pome/Godeps/_workspace/src/github.com/jmoiron/sqlx"
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
	"USERINFO":        []string{"userid", "USERID", "Badgenumber", "name", "Gender", "BIRTHDAY", "CardNo", "lastname", "identitycard", "bankcode1", },
}

type Config struct {
	dbAdapter      string
	connString     string
	sqlQuery       string
	limitQuery     string
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
		chipName:   "חדר אוכל קטן", },
	"רחבה סעודות":{
		controller: "ארועים",
		chipName:   "רחבת סעודות 1", },
	"רמפה":{
		controller: "ארועים",
		chipName:   "ק.0 מדרגות גימס סעודות", },
	"סעודות עולם אירועים מערב":{
		controller: "חדר אמנון פעיל קומה ",
		chipName:   "ק.0 ארועים מערב סעודות", },
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

func loadPrices(db *sqlx.DB, cfg *Config) {
	printCommand("loadPrices", "")
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
			income TEXT,
			meal TEXT,
			start TEXT,
			finish TEXT,
			controller TEXT,
			chip_name TEXT,
			price INTEGER,
			vegetarian INTEGER
		)`)
	sheet := xlFile.Sheets[0]
	for _, row := range sheet.Rows {
		cells := row.Cells
		//fmt.Println(cells)
		if len(cells) < 11 || cells[1].Value == "" {
			continue
		}
		if _, err = strconv.Atoi(cells[1].Value); err != nil {
			continue
		}

		day, _ := strconv.Atoi(cells[1].Value)
		dow, ok := daysMap[cells[2].Value]
		if !ok {
			fmt.Errorf("### Unknown DOW: %d\n", cells[2].Value)
			os.Exit(-1)
		}
		s := strings.Split(cells[5].Value, "-")
		start, end := strings.Replace(s[0], ".", ":", -1), strings.Replace(s[1], ".", ":", -1)
		priceChip, _ := strconv.Atoi(cells[7].Value)
		priceVegetarian, _ := strconv.Atoi(cells[9].Value)
		label := strings.TrimSpace(strings.Replace(cells[4].Value, "'", "׳", -1))
		income := strings.TrimSpace(strings.Replace(cells[3].Value, "'", "׳", -1))
		for _, chipName := range strings.Split(cells[6].Value, ";") {
			controller, ok := controllerMap[strings.TrimSpace(chipName)]
			if !ok {
				fmt.Errorf("### Unable to find chip: %s\n", chipName)
				os.Exit(-1)
			}
			query := fmt.Sprintf(`
				INSERT INTO prices (day, dow, income, meal, start, finish, controller, chip_name, price, vegetarian)
				VALUES(%d, %d, '%s', '%s', '%s', '%s', '%s', '%s', %d, %d);
			`, day, dow, income, label, start, end, controller.controller, controller.chipName, priceChip, priceVegetarian)
			if _, err = db.Exec(query); err != nil {
				log.Fatal(err)
				os.Exit(-1)
			}
		}
	}
}

func getConfig() *Config {

	thisYear, thisMonth, _ := time.Now().Date()

	var doNotLoadMdb = flag.Bool("x", false, "Do not reload MDB file, use the existing data")
	var onlyLoadMdb = flag.Bool("X", false, "ONLY reload MDB file")
	var mdbPath = flag.String("m", "ZKAccess.mdb", "Path to an MDB file")
	var outputPath = flag.String("o", "", "** mandatory ** Output file path")
	var month = flag.Int("d", int(thisMonth - 1), "Month to create report for")
	var year = flag.Int("y", int(thisYear), "Year to create report for")
	var pUser = flag.String("u", "postgres", "Postgres User")
	var pPassword = flag.String("p", "postgres", "Postgres Password")
	var pHost = flag.String("h", "localhost", "Postgres host")
	var makor = flag.String("i", "", "File with prices")
	var noMakor = flag.Bool("I", false, "Do not load file with prices")
	var doNotCalculate = flag.Bool("C", false, "Do not perform calculations")

	if *month == 0 {
		*month = 12
	}
	if *month == 12 {
		*year --
	}

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

	limitQ := fmt.Sprintf(`
		DELETE FROM acc_monitor_log WHERE time < '%d-%02d-01 00:00:00' OR time >= '%d-%02d-01 00:00:00';
	`, *year, *month, *year, *month+1)

	query := fmt.Sprintf(`

		WITH _in AS (
			SELECT * FROM all_records WHERE meal IN ('סעודת בוקר', 'סעודת צהריים')
		),
		_not_in AS (
			SELECT DISTINCT * FROM all_records WHERE meal NOT IN ('סעודת בוקר', 'סעודת צהריים')
		),
		meals AS (
			SELECT * FROM _in
			UNION ALL
			SELECT * FROM _not_in
		)
		SELECT  name "שם",
			lastname "שם משפחה",
			'' "טלפון",
			ophone "מס כרטיס בהנה""ח",
			income "הכנסה",
			meal "סוג סעודה",
			'' "פסצקה",
			case when pager='1' then vegetarian else price end "מחיר",
			day || '.%02d.%d' "תאריך",
			case when pager='1' then 'כן' else '' end "צמחוני",
			pin || ',' || card_no "badge/tag",
			'' "הערות",
			CASE
				WHEN device_name = 'ארועים' THEN
					CASE
						WHEN event_point_name = 'חדר אוכל קטן' THEN 'חדר אוכל'
						WHEN event_point_name = 'רחבת סעודות 1' THEN 'רחבה סעודות'
						WHEN event_point_name = 'ק.0 מדרגות ג''מס סעודות' THEN 'רמפה'
					END
				WHEN device_name = 'חדר אמנון פעיל קומה ' THEN
					CASE WHEN event_point_name = 'ק.0 ארועים מערב סעודות' THEN 'סעודות עולם אירועים מערב'
					END
			END "קולט"
		FROM meals
		WHERE meal IS NOT NULL
		ORDER BY "מס כרטיס בהנה""ח", "שם משפחה", "שם", day;
		`, *month, *year)

	return &Config{
		dbAdapter:      "postgres",
		connString:     fmt.Sprintf("postgres://%s%s@%s/zkaccess", *pUser, *pPassword, *pHost),
		sqlQuery:       query,
		limitQuery:     limitQ,
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

func loadDB(dbpath string, tables dbStruct, db *sqlx.DB, cfg *Config) (err error) {
	printCommand("LoadDB", "")

	for table := range tables {
		dbTable := strings.ToLower(table)

		db.Exec("ALTER SEQUENCE " + dbTable + "_" + tables[table][0] + "_seq RESTART WITH 1;")

		if err = runCommand("Create table "+table,
			"mdb-schema --drop-table -T "+table+" "+dbpath+" postgres | fgrep -v 'ADD CONSTRAINT' | tr '[:upper:]' '[:lower:]' | psql -U postgres -d zkaccess"); err != nil {
			return
		}
		if err = runCommand("Import table "+table,
			"mdb-export -H -q \\\" -D '%Y-%m-%d %H:%M:%S' "+dbpath+" "+table+"| psql -U postgres -d zkaccess -c 'COPY "+dbTable+" FROM STDIN CSV'"); err != nil {
			return
		}
	}

	db.Exec(cfg.limitQuery)
	// keep only records that belong to our events
	db.Exec("DROP TABLE IF EXISTS events;")
	into_events := `
		  SELECT pin, card_no, device_name, event_point_name,
			extract(dow from time)::integer AS dow,
			extract(day from time)::integer AS day,
			to_char(time, 'HH24:MI') AS hm
		  INTO events
		  FROM acc_monitor_log
		  WHERE false
	`
	for _, v := range controllerMap {
		into_events += fmt.Sprintf(` OR (device_name = '%s' AND event_point_name = '%s')`, v.controller, v.chipName)
	}
	db.Exec(into_events)

	for_kolia, _ := db.Queryx(`
		SELECT day, hm, card_no, device_name, event_point_name
		FROM events e
		LEFT OUTER JOIN userinfo u ON u.badgenumber = e.pin OR u.cardno = e.card_no
		WHERE (coalesce(u.name, '') || ' ' || coalesce(u.lastname, '')) = ' '
	`)
	file := xlsx.NewFile()
	defer file.Save(cfg.outputPath + "/for_kolia" + strconv.Itoa(cfg.month) + "-" + strconv.Itoa(cfg.year) + ".xlsx")
	writeSheet(for_kolia, file)

	db.Exec(`
		DROP TABLE IF EXISTS all_records;
		SELECT e.*, u.name, u.lastname, u.ophone, u.fphone, u.pager, p.meal, p.price, p.vegetarian, p.income
		INTO all_records
		FROM events e
		LEFT OUTER JOIN userinfo u ON u.badgenumber = e.pin OR u.cardno = e.card_no
		LEFT OUTER JOIN prices p ON e.dow = p.dow AND e.day = p.day AND e.hm BETWEEN p.start AND p.finish
		AND e.device_name = p.controller AND e.event_point_name = p.chip_name
		WHERE (coalesce(u.name, '') || ' ' || coalesce(u.lastname, '')) != ' ';
		-- ignore oraat keva
		DELETE FROM all_records a WHERE a.dow = 5 AND a.fphone = '2' AND a.hm > '15:00';
		DELETE FROM all_records a WHERE a.dow = 6 AND a.fphone = '2';
	`)

	return
}

func runCommand(description string, command string) (err error) {
	var out []byte

	printCommand(description, command)
	cmd := exec.Command("sh", "-c", command)
	if out, err = cmd.CombinedOutput(); err != nil {
		printError(err)
		printOutput(out)
	}
	return
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
		WITH _in AS (
			SELECT * FROM all_records WHERE meal IN ('סעודת בוקר', 'סעודת צהריים')
		),
		_not_in AS (
			SELECT DISTINCT * FROM all_records WHERE meal NOT IN ('סעודת בוקר', 'סעודת צהריים')
		),
		meals AS (
			SELECT * FROM _in
			UNION ALL
			SELECT * FROM _not_in
		),
		run_int AS (
			SELECT meal, day, count(1), SUM(CASE WHEN pager = '1' THEN vegetarian ELSE price END)
			FROM meals
			GROUP BY meal, day
			ORDER BY meal, day ASC
		),
		run_tot AS (
			SELECT meal, 0 AS day, count(1), SUM(CASE WHEN pager = '1' THEN vegetarian ELSE price END)
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
	printCommand("LcalculateMoney", "")
	results, err := db.Queryx(cfg.sqlQuery)
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
		ELSE device_name || '--' || event_point_name
	END "device"
	FROM acc_monitor_log
	WHERE device_name IN ('ארועים', 'חדר אמנון פעיל קומה ') AND
	      event_point_name IN ('חדר אוכל קטן', 'רחבת סעודות 1', 'ק.0 מדרגות גימס סעודות', 'ק.0 ארועים מערב סעודות')
	)
	SELECT count(device), device, device_name, event_point_name
	FROM a
	GROUP BY device, device_name, event_point_name`)
	if err != nil {
		log.Fatal(err)
		os.Exit(-1)
	}

	file := xlsx.NewFile()
	defer file.Save(cfg.outputPath + "/chip_reader" + strconv.Itoa(cfg.month) + "-" + strconv.Itoa(cfg.year) + ".xlsx")

	writeSheet(results, file)

	// Price per "מס כרטיס בהנה""ח"
	results, err = db.Queryx(`
			WITH _in AS (
				SELECT * FROM all_records WHERE meal IN ('סעודת בוקר', 'סעודת צהריים')
			),
			_not_in AS (
				SELECT DISTINCT * FROM all_records WHERE meal NOT IN ('סעודת בוקר', 'סעודת צהריים')
			),
			meals AS (
				SELECT * FROM _in
				UNION ALL
				SELECT * FROM _not_in
			)
	select ophone "מס כרטיס בהנה""ח", sum(case when pager = '1' then vegetarian else price end) "מחיר"
	from meals
	GROUP BY ophone
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
