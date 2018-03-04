Sync controllers: "ארועים" and "חדר אמנון פעיל קומה " ;rmadmin; doorkeys; admin/admin
Copy \\bb-bakaradb\ZKAccess3.5\ZKAccess.mdb
Run ```go build bb-bakara.go``` to create linux executable
prices: https://docs.google.com/spreadsheets/d/1UvhtwUb9nl-K_WK_v2dVNbDPScwjE5FwetTEOCeA_7Q/edit#gid=1907896725
go build bb-bakara.go
./bb-bakara -d 01 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-01.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-01
./bb-bakara -d 02 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-02.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-02
./bb-bakara -d 03 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-03.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-03
./bb-bakara -d 04 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-04.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-04
./bb-bakara -d 05 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-05.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-05
./bb-bakara -d 06 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-06.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-06
./bb-bakara -d 07 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-07.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-07
./bb-bakara -x -d 8 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-08.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-08
./bb-bakara -d 9 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-09.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-09
./bb-bakara -d 10 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-10.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-10
./bb-bakara -d 11 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-11.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-11
./bb-bakara -d 12 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-12-weekend.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-12-weekend
./bb-bakara -d 12 -y 2017 -i /media/sf_projects/bb-bakaradb/reports/prices2017-12.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2017-12
./bb-bakara -d 01 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-01-13.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-01-13
./bb-bakara -d 01 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-01-26.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-01-26
./bb-bakara -d 02 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-02-02.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-02-02
./bb-bakara -x -d 01 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-01.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-01
./bb-bakara -d 02 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-02-09.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-02-09
./bb-bakara -xC -d 02 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-02.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-02



-- for Nahari
select name "שם", lastname "שם משפחה", ophone "מס' כרטיס בהנה""ח",
	case when pager = '1' then 'כן' else '' end "צמחוני",
	case when fphone = '2' then 'כן' else '' end "שבט",
	cardno "צ'יפ"
from userinfo


=INDEX(Kolia!G:G,MATCH(A2,Kolia!A:A,0), 0)