FROM mysql:8.0.23

MAINTAINER caneserpente

ENV MYSQL_DATABASE=mempoi \
    MYSQL_ROOT_PASSWORD=mempassword

ADD data/test_dump.sql /docker-entrypoint-initdb.d

EXPOSE 3306
