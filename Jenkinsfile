pipeline {
    agent any

    triggers {
        cron('0 5 * * *') // Runs daily at 5:00 AM server time
    }

    environment {
      //  JAVA_HOME = '/path/to/your/java'    // Set this on Jenkins machine
        PATH = "${JAVA_HOME}/bin:${env.PATH}"
        CHROME_DRIVER = '/path/to/chromedriver' // Optional: use if needed
    }

    stages {
        stage('Checkout Code') {
            steps {
                echo 'Cloning GitHub repository...'
                git branch: 'master', url: 'https://github.com/ansul9052/AmazonData.git'
            }
        }

        stage('Build') {
            steps {
                echo 'Compiling Java sources...'
                // Adjust classpath to include your dependencies; 'libs/*' expects you to have JARs there or adjust accordingly
                sh 'javac -cp "libs/*" src/test/java/Amazon/in/*.java'
            }
        }

        stage('Run Cab_Data') {
            steps {
                echo 'Running Cab_Data script...'
                sh 'java -cp "libs/*:src/test/java" Amazon.in.Cab_Data'
            }
        }

        stage('Run EHS') {
            steps {
                echo 'Running EHS script...'
                sh 'java -cp "libs/*:src/test/java" Amazon.in.EHS'
            }
        }
    }
}
