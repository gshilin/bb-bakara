Sync controllers: "ארועים" and "חדר אמנון פעיל קומה " ;rmadmin; doorkeys; admin/eA3Z5xJe
Copy \\bb-bakaradb\ZKAccess3.52\ZKAccess.mdb (bbdomain\falcon)
Run ```go build bb-bakara.go && strip bb-bakara && upx -9 bb-bakara``` to create linux executable
prices: https://docs.google.com/spreadsheets/d/1UvhtwUb9nl-K_WK_v2dVNbDPScwjE5FwetTEOCeA_7Q/edit#gid=1907896725
go build bb-bakara.go
./bb-bakara -d 06 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-06.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-06
./bb-bakara -d 07 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-07.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-07
./bb-bakara -d 8 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-08.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-08
./bb-bakara -d 9 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-09.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-09
./bb-bakara -d 10 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-10.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-10
./bb-bakara -d 11 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-11.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-11
./bb-bakara -d 12 -y 2018 -i /media/sf_projects/bb-bakaradb/reports/prices2018-12.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2018-12
./bb-bakara -d 1 -y 2019 -i /media/sf_projects/bb-bakaradb/reports/prices2019-01.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2019-01
./bb-bakara -d 2 -y 2019 -i /media/sf_projects/bb-bakaradb/reports/prices2019-02.xlsx -m /media/sf_projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_projects/bb-bakaradb/reports/2019-02
./bb-bakara -d 3 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-03.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-03
./bb-bakara -d 3 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-03a.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-03a
./bb-bakara -d 4 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-04.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-04
./bb-bakara -d 5 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-05.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-05
./bb-bakara -d 6 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-06.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-06
./bb-bakara -d 7 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-07.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-07
./bb-bakara -d 8 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-08.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/Access.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-08
./bb-bakara -d 9 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-09.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-09
./bb-bakara -d 10 -y 2019 -i /media/sf_D_DRIVE/projects/bb-bakaradb/reports/prices2019-10.xlsx -m /media/sf_D_DRIVE/projects/bb-bakaradb/ZKAccess.mdb -o /media/sf_D_DRIVE/projects/bb-bakaradb/reports/2019-10



-- for Nahari
select name "שם", lastname "שם משפחה", ophone "מס' כרטיס בהנה""ח",
	case when pager = '1' then 'כן' else '' end "צמחוני",
	case when fphone = '2' then 'כן' else '' end "שבט",
	cardno "צ'יפ"
from userinfo


=INDEX(Kolia!G:G,MATCH(A2,Kolia!A:A,0), 0)


acc_levelset 39
acc_levelset_door_group
    id 6392
    accdoor_id 58
    accdoor_no_exp 2
    accdoor_device_id 17
    level_timeseg_id 10

acc_levelset_emp
    acc_levelset_id = 39
    employee_id

           SELECT userinfo.name, userinfo.lastname
           from userinfo
           where userinfo.userid in (select employee_id from acc_levelset_emp where acclevelset_id = 39)
