# Use a Node.js base image to run the npm commands
FROM node:24-alpine AS build_stage

# Set the working directory inside the container
WORKDIR /app

RUN apk add --no-cache bash

# Copy package.json and package-lock.json and install dependencies
# This step is cached and only re-run if these files change
COPY package*.json ./
RUN npm install

# Copy the rest of the application source code
COPY . .

# Run the build command to generate the production files
# The output will be in the 'dist' directory
RUN npm run clean && npm run dist

# -----------------------------------------------------------------------------

# Use the official Tomcat image for the final production environment
FROM tomcat:11.0.11-jre17-temurin AS final_stage

# Copy the built files from the 'dist/public' directory of the build stage
# into the correct location in the Tomcat webapps directory.
COPY --from=build_stage /app/dist/public /usr/local/tomcat/webapps/ROOT/
COPY --from=build_stage /app/dist/server.xml /usr/local/tomcat/conf/server.xml

# Clean up any default Tomcat files if they are not needed
RUN rm -rf /usr/local/tomcat/webapps/ROOT/WEB-INF/ \
    /usr/local/tomcat/webapps/ROOT/*.jsp

# Expose the default Tomcat port
EXPOSE 8080 3000

ARG KEYSTORE_PASS
RUN if [ "$KEYSTORE_PASS" != "" ]; then \
      echo "Generating SSL certificates..."; \
      keytool -genkey -alias tomcat -keyalg RSA -keystore /usr/local/tomcat/conf/localhost-rsa.jks \
        -storetype JKS -storepass ${KEYSTORE_PASS} -keypass ${KEYSTORE_PASS} -validity 365 \
        -dname "CN=localhost, OU=IT, O=MyCompany, L=MyCity, ST=MyState, C=US"; \
      chmod a+rx /usr/local/tomcat/conf/localhost-rsa.jks; \
    fi

ARG account=tomcat
# Combined addgroup and adduser commands
RUN addgroup --system ${account} && adduser ${account} --system &&\
     adduser ${account} ${account}

USER ${account}

# Use the inherited CMD from the base image to start the server