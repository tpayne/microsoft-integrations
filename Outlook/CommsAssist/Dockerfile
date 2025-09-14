FROM tomcat:10.1.20-jre17-temurin AS base

WORKDIR /usr/local/tomcat

COPY dist/public /usr/local/tomcat/webapps/ROOT/

RUN rm -rf /usr/local/tomcat/webapps/ROOT/WEB-INF/ \
    /usr/local/tomcat/webapps/ROOT/*.jsp

EXPOSE 8080

ARG account=tomcat
RUN addgroup --system ${account} && adduser ${account} --system &&\
     adduser ${account} ${account}
USER ${account}
