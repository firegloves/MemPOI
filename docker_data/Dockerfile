FROM mysql:8-debian

ENV MYSQL_DATABASE=mempoi \
    MYSQL_ROOT_PASSWORD=mempassword

COPY data/test_dump.sql /docker-entrypoint-initdb.d

EXPOSE 3306
