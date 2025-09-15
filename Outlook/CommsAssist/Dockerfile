# Use a Node.js base image to run the npm commands
FROM node:18-alpine AS build_stage

# Set the working directory inside the container
WORKDIR /app

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
FROM tomcat:10.1.20-jre17-temurin AS final_stage

# Copy the built files from the 'dist/public' directory of the build stage
# into the correct location in the Tomcat webapps directory.
COPY --from=build_stage /app/dist/public /usr/local/tomcat/webapps/ROOT/

# Clean up any default Tomcat files if they are not needed
RUN rm -rf /usr/local/tomcat/webapps/ROOT/WEB-INF/ \
    /usr/local/tomcat/webapps/ROOT/*.jsp

# Expose the default Tomcat port
EXPOSE 8080

ARG account=tomcat
# Combined addgroup and adduser commands
RUN addgroup --system ${account} && adduser ${account} --system &&\
     adduser ${account} ${account}
USER ${account}

# Use the inherited CMD from the base image to start the server