# Deployment / Provisioning mit nodejs

##install
npm install
typings install

##build
tsc || tsc -w (mit watch)

#parameters
-f : <configfile>
-p : <password>
-x : flag -> proxy

##run
node /dist/deploy -f /config/democonfig.json
