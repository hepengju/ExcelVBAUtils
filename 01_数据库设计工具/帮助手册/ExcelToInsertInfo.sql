Rem /*
Rem 工具: Excel_To_Insert 
Rem 
Rem 作者: 何鹏举
Rem 版本: V1.4
Rem 改进: 1.所有模块重新思索考虑并重写,使得更加规范化(变量命名,注释添加,逻辑调整,界面调整等)
Rem       2.增加MySQL,PostgreSQL等支持的多值插入格式的支持: Insert into (...) values(...),(...),(...)的格式
Rem       3.批次提交考虑更多数据库的事务写法,并兼容最新的多值插入格式
Rem 说明: 如下取10条数据,简要列出功能
Rem */
Rem 
Rem /*0.列类型说明: 第一行的底纹不是白色或字体不是黑色,当做数据,否则当做字符(日期型对于各个数据库差异较大,过于复杂,不再考虑)*/
Rem /*1.生成语句选项 */
Rem     /*1.1数据列名(否): 此时插入列名无论是否均不生效*/
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem     /*1.2数据列名(是),插入列名(否)*/
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem     /*1.3数据列名(是),插入列名(是)*/
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem /*2.保存文件选项(自动保存在桌面)*/
Rem     /*2.1清空源表(是): 写入Truncate语句*/
Rem         -- 清空表数据
Rem         TRUNCATE TABLE tableName;
Rem 
Rem         -- 开始插入数据
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         -- 所有数据(10条)插入完毕
Rem     /*2.2多值插入(是): 采用Insert into (...) values(...),(...),(...)的格式*/
Rem         -- 开始插入数据
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) 
Rem              VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem           ;
Rem         -- 所有数据(10条)插入完毕
Rem     /*2.3批次提交(5): 分5条数据提交一次,其中开启事务和提交事务的语句可以另外选择,默认才用Oracle的格式即开启事务为"空白",提交事务为"COMMIT;"*/
Rem         
Rem         /*2.3.1普通的批次提交 空白,COMMIT;*/
Rem         -- 开始插入数据
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         COMMIT;
Rem         -- 第5条数据完毕
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         COMMIT;
Rem         -- 第10条数据完毕
Rem         -- 所有数据(10条)插入完毕
Rem         
Rem         /*2.3.1普通的批次提交 BEGIN,END*/
Rem         -- 开始插入数据
Rem         BEGIN
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         END
Rem         -- 第5条数据完毕
Rem         BEGIN
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512);
Rem         END
Rem         -- 第10条数据完毕
Rem         -- 所有数据(10条)插入完毕
Rem         
Rem         /*2.3.3多值插入的批次提交,由于本身语句就会较少,不采用开启和提交事务语句,只是按照批次分割*/
Rem         -- 开始插入数据
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) 
Rem              VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem           ;
Rem         -- 第5条数据完毕
Rem         INSERT INTO tableName (FSD, KHRQ, DQRQ, BZ, SXED, GXED, DBFS, DKLL) 
Rem              VALUES ('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem                    ,('331100', '20120329', '20120529', 'CNY', 70000, 70000, '1', 512)
Rem           ;
Rem         -- 第10条数据完毕
Rem         -- 所有数据(10条)插入完毕