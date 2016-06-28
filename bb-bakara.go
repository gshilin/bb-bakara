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
)

type persons_type map[string]int

type dbStruct map[string][]string

var tables = dbStruct{
	//"acc_door": []string{"id", "device_id", "door_no", "door_name"},
	//"acc_levelset": []string{"id", "level_name"},
	//"acc_levelset_door_group": []string{"id", "acclevelset_id", "accdoor_id", "accdoor_no_exp", "accdoor_device_id"},
	//"acc_levelset_emp": []string{"id", "acclevelset_id", "employee_id"},
	"acc_monitor_log": []string{"id", "time", "pin", "card_no", "device_id", "device_sn", "device_name", "event_point_name"},
	//"DEPARTMENTS": []string{"id", "DEPTNAME", "SUPDEPTID"},
	//"Machines": []string{"id", "MachineAlias", "IP", "SerialPort", "Port", "Baudrate", "usercount", "FirmwareVersion",
	//	"sn", "device_name", "subnet_mask", "gateway", "area_id", "acpanel_type"},
	//"personnel_area": []string{"id", "areaid", "areaname", "parent_id"},
	//"personnel_issuecard": []string{"id", "create_time", "UserID_id", "cardno", },
	"USERINFO": []string{"id", "USERID", "Badgenumber", "name", "Gender", "BIRTHDAY", "CardNo", "lastname", "identitycard", "bankcode1", },
}

type Config struct {
	dbAdapter    string
	connString   string
	sqlQuery     string
	outputFile   string
	mdbPath      string
	doNotLoadMdb bool
}

func main() {
	cfg := getConfig()

	if !cfg.doNotLoadMdb {
		loadDB(cfg.mdbPath, tables)
	}

	calculateMoney(cfg)
}

func getConfig() *Config {

	thisYear, thisMonth, _ := time.Now().Date()
	var doNotLoadMdb = flag.Bool("x", false, "Path to an MDB file")
	var mdbPath = flag.String("m", "ZKAccess.mdb", "Path to an MDB file")
	var outputPath = flag.String("o", "", "Output file path and name ** mandatory **")
	var month = flag.Int("d", int(thisMonth - 1), "Month to create report for")
	var year = flag.Int("y", int(thisYear - 1), "Year to create report for")
	var pUser = flag.String("u", "postgres", "Postgres User")
	var pPassword = flag.String("p", "postgres", "Postgres Password")
	var pHost = flag.String("h", "localhost", "Postgres host")

	flag.Parse()

	if *outputPath == "" {
		flag.PrintDefaults()
		log.Fatal("Please supply all mandatory parameters")
	}
	if *month <= 0 || *month > 12 {
		log.Fatal("Bad month number", *month)
	}

	if (*pPassword == "") {
		*pPassword = ""
	} else {
		*pPassword = ":" + *pPassword
	}

	query := fmt.Sprintf(`
CREATE EXTENSION IF NOT EXISTS tablefunc;

WITH crosstab AS (
	SELECT * FROM crosstab(
		$$ SELECT * FROM (
			WITH records AS (
				SELECT
					pin || ',' || card_no AS name,
					time, device_name, event_point_name,
					extract(dow from time)::integer AS dow,
					to_char(time, 'HH24:MI') AS hm
				FROM acc_monitor_log a
				WHERE   time >= '%d-%02d-01 00:00:00' AND time < '%d-%02d-01 00:00:00'
					and (
						(device_name = 'חדר אמנון פעיל קומה ' and event_point_name = 'ק.1 טכנאי מחשב חדר אמנון')
						or
						(device_name = 'גלריה וכניסה' and event_point_name = 'ארוחת צהרים חדר אוכל')
					)
			)
			SELECT name, 'breakfast' AS title, count(1) * 3 AS shekel FROM records
			WHERE dow BETWEEN 0 AND 4 AND hm BETWEEN '02:00' AND '11:30' GROUP BY name
				UNION
			SELECT name, 'dinner' AS title, count(1) * 14 AS shekel FROM records
			WHERE dow BETWEEN 0 AND 4 AND hm BETWEEN '11:31' AND '18:00' GROUP BY name
				UNION
			SELECT name, 'houmus' AS title, count(1) * 5 AS shekel FROM records
			WHERE dow = 5 AND hm BETWEEN '02:00' AND '11:00' GROUP BY name
		) meals ORDER BY name $$,
		$$ VALUES ('breakfast'::text), ('dinner'::text), ('houmus'::text) $$
	) AS (name TEXT, breakfast BIGINT, dinner BIGINT, houmus BIGINT)
), a AS (
	SELECT c.name, coalesce(u.name, '') || ' ' || coalesce(u.lastname, '') full_name, u.ophone, breakfast, dinner, houmus, SUM(coalesce(breakfast, 0) + coalesce(dinner, 0) + coalesce(houmus, 0)) AS total
	FROM crosstab c
	LEFT OUTER JOIN userinfo u ON (u.badgenumber = split_part(c.name, ',', 1))
	GROUP BY c.name, breakfast, dinner, houmus, ophone, full_name
), by_badge AS (
	SELECT * FROM a WHERE full_name != ' '
), by_tag AS (
	SELECT a.name, coalesce(u.name, '') || ' ' || coalesce(u.lastname, '') full_name, a.ophone, breakfast, dinner, houmus, total
	FROM a
	LEFT OUTER JOIN userinfo u ON u.cardno = split_part(a.name, ',', 2)
	WHERE full_name = ' '
), final AS (
	SELECT * FROM by_badge UNION SELECT * FROM by_tag
)

SELECT ophone "מס כרטיס בהנה""ח", name "badge/tag", full_name "שם מלא", breakfast "סעודות בוקר", dinner "סעודות צהריים", houmus "סעודת חומוס", total "סה""כ"
FROM (
	 SELECT 0 ordinal, * FROM final
	UNION
	 SELECT 1 ordinal, '', '', 'Grand Total', SUM(coalesce(breakfast, 0)), SUM(coalesce(dinner, 0)), SUM(coalesce(houmus, 0)), SUM(coalesce(total, 0)) FROM final
) t
ORDER BY ordinal ASC, "badge/tag" ASC;
		`, *year, *month, *year, *month + 1)

	return &Config{
		dbAdapter: "postgres",
		connString: fmt.Sprintf("postgres://%s%s@%s/zkaccess", *pUser, *pPassword, *pHost),
		sqlQuery: query,
		outputFile: *outputPath,
		mdbPath: *mdbPath,
		doNotLoadMdb: *doNotLoadMdb,
	}
}

func loadDB(dbpath string, tables dbStruct) (err error) {

	for table := range tables {
		dbTable := strings.ToLower(table)

		if err = runCommand("Create table " + table,
			"mdb-schema --drop-table -T " + table + " " + dbpath +
				" postgres | fgrep -v 'ADD CONSTRAINT' | tr '[:upper:]' '[:lower:]' | psql -U postgres -d zkaccess"); err != nil {
			return
		}
		if err = runCommand("Import table " + table,
			"mdb-export -H -q \\\" -D '%Y-%m-%d %H:%M:%S' " + dbpath + " " + table +
				"| psql -U postgres -d zkaccess -c 'COPY " + dbTable + " FROM STDIN CSV'"); err != nil {
			return
		}
	}

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

func calculateMoney(cfg *Config) {

	db, err := sqlx.Open(cfg.dbAdapter, cfg.connString)
	if err != nil {
		log.Fatal(err)
	}

	results, err := db.Queryx(cfg.sqlQuery)

	if err != nil {
		log.Fatal(err)
	}
	file := xlsx.NewFile()
	defer file.Save(cfg.outputFile)

	sheet, err := file.AddSheet("Sheet1") // col bestFit="1"
	if err != nil {
		log.Fatal(err)
	}
	row := sheet.AddRow()
	columns, _ := results.Columns()
	for _, h := range columns { //[]string{"מס כרטיס בהנה\"ח", "badge id, tag id", "שם מלא", "סעודות בוקר", "סעודות צהריים", "סעודת חומוס", "סה\"כ"} {
		cell := row.AddCell()
		cell.Value = h
	}
	for results.Next() {
		xRow := sheet.AddRow()
		row, err := results.SliceScan()
		if err != nil {
			log.Fatal(err)
		}

		for _, col := range row {
			//log.Print(reflect.TypeOf(col))
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
}
