namespace Mergedt.DB
{
    public class SqlList
    {
        //根据SQLID返回对应的SQL语句  
        private string _result;

        /// <summary>
        /// 初始化获取配方号以及来源ID
        /// </summary>
        /// <returns></returns>
        public string GetSourcedt()
        {
            _result = @"SELECT a.ColorCode,b.RegionId FROM dbo.InnerColor a
                        INNER JOIN dbo.RegionColor b ON a.InnerColorId=b.InnerColorId
                        ORDER BY a.ColorCode";
            return _result;
        }

        /// <summary>
        /// 已没有使用
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        public string GetIncloudSource(string code)
        {
            _result = $@"
                            DECLARE
	                            @count INT,
	                            @id INT,
	                            @name NVARCHAR(100),
	                            @result INT;
                            BEGIN
	                            SELECT @count=COUNT(*),@id=a.InnerColorId
	                            FROM dbo.InnerColor a
	                            INNER JOIN dbo.RegionColor b ON a.InnerColorId=b.InnerColorId 
	                            INNER JOIN dbo.Relations c ON b.RegionId=c.RelationsId
	                            WHERE a.ColorCode='{code}'
	                            GROUP BY a.InnerColorId

	                            SELECT @name=b.RelationsName
	                            FROM dbo.RegionColor a
	                            INNER JOIN dbo.Relations b ON a.RegionId=b.RelationsId
	                            WHERE a.InnerColorId=@id

	                            IF(@count=1 AND @name='CHINA')
	                                BEGIN
		                                SET @result='1'
                                    END
	                            ELSE
	                                BEGIN
	                                    SET @result='0'
	                                END       

	                            SELECT @result
                            END
                        ";

            return _result;
        }


    }
}
