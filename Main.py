__author__ = 'yanglikun'
# 首先在config.properties配置数据库信息
from mysql2doc.Document import Word
from mysql2doc.config import dbConfig

Word.createFile(dbConfig.databaseName)
