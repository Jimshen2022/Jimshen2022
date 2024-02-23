CREATE DATABASE MICK

CREATE TABLE Greatests
(Key1    CHAR(10)    NOT NULL,
 x1      INTEGER     ,
 y1      INTEGER     ,
 z1      INTEGER     ,
 PRIMARY KEY (Key1)); 
 
 ALTER TABLE Greatests DROP COLUMN z1;
 ALTER TABLE Greatests ADD COLUMN z1 INTEGER;
  
 
 SELECT *
 FROM Greatests
 
 BEGIN TRANSACTION;
 INSERT INTO Greatests VALUES ('A',1,2,3);
 INSERT INTO Greatests VALUES ('B',5,5,2);
 INSERT INTO Greatests VALUES ('C',4,7,1);
 INSERT INTO Greatests VALUES ('D',3,3,8);
 COMMIT;
 

 DROP TABLE greatests;
 
 
SELECT Key1,
  (CASE WHEN Qty>Qty2 THEN Qty ELSE Qty2 END) AS qty3
 
 FROM (  
	SELECT Key1, 
	MAX(CASE WHEN x1>y1 THEN x1 ELSE y1 END) AS Qty, MAX(CASE WHEN y1>z1 THEN y1 ELSE z1 END) AS Qty2
	FROM greatests
	GROUP BY key1) AS jim
 GROUP BY Key1 
 
 
 
 # 在UPDATE语句里进行条件分支
 
DROP TABLE Salary;

 
 CREATE TABLE Salaries(
	NAME VARCHAR(20),
	salary INT
	);
	
BEGIN TRANSACTION;
INSERT INTO Salaries VALUES ('相田',300000);
INSERT INTO Salaries VALUES ('神崎',270000);
INSERT INTO Salaries VALUES ('木村',220000);
INSERT INTO Salaries VALUES ('齐藤',290000);
COMMIT;

DELETE FROM Salaries;  -- 删除数据，不删除表
	

CREATE TABLE CourseMaster(
	Course_id INT,
	course_name VARCHAR(20)
	);
	
BEGIN TRANSACTION;
INSERT INTO CourseMaster VALUES ('1','会计入门');
INSERT INTO CourseMaster VALUES ('2','财务知识');
INSERT INTO CourseMaster VALUES ('3','薄记考试');
INSERT INTO CourseMaster VALUES ('4','税务师');
COMMIT;	


SELECT * FROM coursemaster;


DELETE FROM coursemaster  
WHERE course_name = '税务师';

DELETE FROM coursemaster
WHERE course_name LIKE '%师%';


CREATE TABLE OpenCoures(
	MONTH  VARCHAR(20),
	course_id INT
	);
  

UPDATE coursemaster
	SET course_name = '注册会计师'
WHERE course_name = '税务师';


SELECT * FROM coursemaster;

 ALTER TABLE coursemaster ADD COLUMN 分数 INT;
 
ALTER TABLE coursemaster DROP COLUMN 分数;
 
 
 BEGIN TRANSACTION;
 UPDATE coursemaster
	SET 分数 = 99999
	WHERE course_id = 2; 
COMMIT;
 
 BEGIN TRANSACTION;
 UPDATE coursemaster
	SET 分数 = 150
	WHERE course_id = 2;
COMMIT;
	
 
 
 
 
 
 
 
 
 
 