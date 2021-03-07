FROM entando/entando-eap72-clustered-base:6.2.3

COPY target/*.war /opt/eap/standalone/deployments/
RUN $ENTANDO_COMMON_PATH/init-derby-from-war.sh
RUN docker exec -i $container mysql -umagento -pmagento magento <<< "source $file"
