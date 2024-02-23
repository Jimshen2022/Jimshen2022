
--查表的索引
Select *
From QSYS2.SYSINDEXES
LIMIT 5


/*
SQL 约束（Constraints）
SQL 约束用于规定表中的数据规则。

如果存在违反约束的数据行为，行为会被约束终止。

约束可以在创建表时规定（通过 CREATE TABLE 语句），或者在表创建之后规定（通过 ALTER TABLE 语句）。
SQL CREATE TABLE + CONSTRAINT 语法

CREATE TABLE table_name
(
column_name1 data_type(size) constraint_name,
column_name2 data_type(size) constraint_name,
column_name3 data_type(size) constraint_name,
....
);

在 SQL 中，我们有如下约束：

    NOT NULL - 指示某列不能存储 NULL 值。
    UNIQUE - 保证某列的每行必须有唯一的值。
    PRIMARY KEY - NOT NULL 和 UNIQUE 的结合。确保某列（或两个列多个列的结合）有唯一标识，有助于更容易更快速地找到表中的一个特定的记录。
    FOREIGN KEY - 保证一个表中的数据匹配另一个表中的值的参照完整性。
    CHECK - 保证列中的值符合指定的条件。
    DEFAULT - 规定没有给列赋值时的默认值。


*/

-- SQL NOT NULL 约束
/*
NOT NULL 约束强制列不接受 NULL 值。

NOT NULL 约束强制字段始终包含值。这意味着，如果不向字段添加值，就无法插入新记录或者更新记录。

下面的 SQL 强制 "ID" 列、 "LastName" 列以及 "FirstName" 列不接受 NULL 值：
*/
CREATE TABLE Persons (
    ID int NOT NULL,
    LastName varchar(255) NOT NULL,
    FirstName varchar(255) NOT NULL,
    Age int
);

-- 添加 NOT NULL 约束 在一个已创建的表的 "Age" 字段中添加 NOT NULL 约束如下所示：
ALTER TABLE Persons
MODIFY Age int NOT NULL;

-- 删除 NOT NULL 约束在一个已创建的表的 "Age" 字段中删除 NOT NULL 约束如下所示：

ALTER TABLE Persons
MODIFY Age int NULL;


-- SQL UNIQUE 约束
/*
UNIQUE 约束唯一标识数据库表中的每条记录。

UNIQUE 和 PRIMARY KEY 约束均为列或列集合提供了唯一性的保证。

PRIMARY KEY 约束拥有自动定义的 UNIQUE 约束。

请注意，每个表可以有多个 UNIQUE 约束，但是每个表只能有一个 PRIMARY KEY 约束。
*/

-- CREATE TABLE 时的 SQL UNIQUE 约束
-- 下面的 SQL 在 "Persons" 表创建时在 "P_Id" 列上创建 UNIQUE 约束：
-- MYSQL:
CREATE TABLE Persons
(
P_Id int NOT NULL,
LastName varchar(255) NOT NULL,
FirstName varchar(255),
Address varchar(255),
City varchar(255),
UNIQUE (P_Id)
)

-- SQL Server / Oracle / MS Access：
CREATE TABLE Persons
(
P_Id int NOT NULL UNIQUE,
LastName varchar(255) NOT NULL,
FirstName varchar(255),
Address varchar(255),
City varchar(255)
)

-- 如需命名 UNIQUE 约束，并定义多个列的 UNIQUE 约束，请使用下面的 SQL 语法：
-- MySQL / SQL Server / Oracle / MS Access：

CREATE TABLE Persons
(
P_Id int NOT NULL,
LastName varchar(255) NOT NULL,
FirstName varchar(255),
Address varchar(255),
City varchar(255),
CONSTRAINT uc_PersonID UNIQUE (P_Id,LastName)
)

-- 当表已被创建时，如需在 "P_Id" 列创建 UNIQUE 约束，请使用下面的 SQL：
ALTER TABLE Persons
ADD UNIQUE (P_Id)

-- 如需命名 UNIQUE 约束，并定义多个列的 UNIQUE 约束，请使用下面的 SQL 语法：
ALTER TABLE Persons
ADD CONSTRAINT uc_PersonID UNIQUE (P_Id,LastName)


-- 撤销 UNIQUE 约束
-- mysql
ALTER TABLE Persons
DROP INDEX uc_PersonID

-- SQL Server / Oracle / MS Access：
ALTER TABLE Persons
DROP CONSTRAINT uc_PersonID



-- SQL PRIMARY KEY 约束
/*
PRIMARY KEY 约束唯一标识数据库表中的每条记录。

主键必须包含唯一的值。

主键列不能包含 NULL 值。

每个表都应该有一个主键，并且每个表只能有一个主键。
*/

-- 下面的 SQL 在 "Persons" 表创建时在 "P_Id" 列上创建 PRIMARY KEY 约束：
-- Mysql
CREATE TABLE Persons
(
P_Id int NOT NULL,
LastName varchar(255) NOT NULL,
FirstName varchar(255),
Address varchar(255),
City varchar(255),
PRIMARY KEY (P_Id)
)

-- 如需命名 PRIMARY KEY 约束，并定义多个列的 PRIMARY KEY 约束，请使用下面的 SQL 语法：
CREATE TABLE Persons
(
P_Id int NOT NULL,
LastName varchar(255) NOT NULL,
FirstName varchar(255),
Address varchar(255),
City varchar(255),
CONSTRAINT pk_PersonID PRIMARY KEY (P_Id,LastName)
)

-- 当表已被创建时，如需在 "P_Id" 列创建 PRIMARY KEY 约束，请使用下面的 SQL：
ALTER TABLE Persons
ADD PRIMARY KEY (P_Id)

-- 如需命名 PRIMARY KEY 约束，并定义多个列的 PRIMARY KEY 约束，请使用下面的 SQL 语法：
ALTER TABLE Persons
ADD CONSTRAINT pk_PersonID PRIMARY KEY (P_Id,LastName)
-- 注释：如果您使用 ALTER TABLE 语句添加主键，必须把主键列声明为不包含 NULL 值（在表首次创建时）。 


-- 如需撤销 PRIMARY KEY 约束，请使用下面的 SQL：
ALTER TABLE Persons
DROP PRIMARY KEY



-- PRIMARY KEY 约束的实例
CREATE TABLE Persons
(
    Id_P int NOT NULL,
    LastName varchar(255) NOT NULL,
    FirstName varchar(255),
    Address varchar(255),
    City varchar(255),
    PRIMARY KEY (Id_P)  //PRIMARY KEY约束
)
CREATE TABLE Persons
(
    Id_P int NOT NULL PRIMARY KEY,   //PRIMARY KEY约束
    LastName varchar(255) NOT NULL,
    FirstName varchar(255),
    Address varchar(255),
    City varchar(255)
)




-- foreign key 用法

create table if not exists per(
  id bigint auto_increment comment '主键',
  name varchar(20) not null comment '人员姓名',
  work_id bigint not null comment '工作id',
  create_time date default '2021-04-02',
  primary key(id),
  foreign key(work_id) references work(id)
)

create table if not exists work(
  id bigint auto_increment comment '主键',
  name varchar(20) not null comment '工作名称',
  create_time date default '2021-04-02',
  primary key(id)
)


--  SQL 通配符
/*
通配符 	描述
% 	替代 0 个或多个字符
_ 	替代一个字符
[charlist] 	字符列中的任何单一字符
[^charlist]
或
[!charlist] 	不在字符列中的任何单一字符

下面的 SQL 语句选取 name 以 "G"、"F" 或 "s" 开始的所有网站：

*/

-- 下面的 SQL 语句选取 name 以 "G"、"F" 或 "s" 开始的所有网站：

SELECT * FROM Websites
WHERE name REGEXP '^[GFs]'; 


-- 下面的 SQL 语句选取 name 以 A 到 H 字母开头的网站：
SELECT * FROM Websites
WHERE name REGEXP '^[A-H]'; 

-- 下面的 SQL 语句选取 name 不以 A 到 H 字母开头的网站：

SELECT * FROM Websites
WHERE name REGEXP '^[^A-H]'; 


-- SQL JOIN
/*
不同的 SQL JOIN

在我们继续讲解实例之前，我们先列出您可以使用的不同的 SQL JOIN 类型：

    INNER JOIN：如果表中有至少一个匹配，则返回行
    LEFT JOIN：即使右表中没有匹配，也从左表返回所有的行
    RIGHT JOIN：即使左表中没有匹配，也从右表返回所有的行
    FULL JOIN：只要其中一个表中存在匹配，则返回行

*/





















