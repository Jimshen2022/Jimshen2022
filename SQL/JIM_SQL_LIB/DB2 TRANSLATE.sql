一、Oracle中语法：

    Translate（string,from_str,to_str）

    eg:select translate('abcdef','abf','cde') from dual;

    结果：cdcdee（将’abcdef’字符串中的’abf’替换为’cde’）


二、DB2中语法：

    Translate（string, to_str,from_str ）

    eg:select translate('abcdef','abf','cde') from sysibm.sysdummy1;

    结果：ababff（将’abcdef’字符串中的’cde’替换为’abf’）


string：需要处理的字符串

from_str：string字符串中需要转换的字符

to_str：需要转换成的字符

（注：在Oracle和DB2中，Translate方法中参数from_str和to_str的位置正好相反）

Translate函数在string中查找from_str中的字符并将其替换为to_str中的字符(单字符替换)。


三、Translate函数使用场景（以DB2为例）：

1、 校验某字段（手机号码，邮政编码，日期…）是否包含除数字以为的字符：

    eg：select trim(translate(’17112345489asx’,’’,’0123456789’)) from sysibm.sysdummy1;

(trim作用是去除空格)

    显示结果：asx

2、将某字段的数字转换为9，字母转换为X：

    eg:select translate（’XGZ201601’,’9…X…’,’0123456789ABCDEF…’）from sysibm.sysdummy1;

(中间参数为10个9和26个X，后面参数为0-9和A-Z)

    显示结果：XXX999999

3、从一段字符串中提取出字母或者数字：

    select translate（’XGZ201601’,’’,’ABCDEF…’）from sysibm.sysdummy1;

    显示结果：201601


四、写在最后：

1、需要转换的字符(from_str)在需要转换成的字符(to_str)中不存在对应，则转换后被截除 :

    eg：select translate（’abcde’,’12’,’bcde’）from sysibm.sysdummy1;

    显示结果：a12

2、在oracle中转换目的字串(to_str)不能为''，因为''在oracle中被视为空值，因此无法匹配而返回为空值 。但是在DB2中则可以进行匹配:

    eg:(oracle)select translate（’abcde’,’abc’,’’）from dual;

    显示结果：

    eg:(DB2)select translate（’abcde’,’’,’abc’）from sysibm.sysdummy1;

    显示结果：de
————————————————
版权声明：本文为CSDN博主「阿飞哥-Jeffrey」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
原文链接：https://blog.csdn.net/feitianlongfei/article/details/78786600