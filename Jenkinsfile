pipeline {
    agent any

    triggers {
        cron('0 5 * * *') // Run daily at 5 AM
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
                echo 'Building the project using Maven...'
                bat 'mvn clean compile'
            }
        }

        stage('Run Cab_Data') {
            steps {
                echo 'Running Cab_Data class...'
                bat 'mvn exec:java -Dexec.mainClass=Amazon.in.Cab_Data'
            }
        }

        stage('Run EHS') {
            steps {
                echo 'Running EHS class...'
                bat 'mvn exec:java -Dexec.mainClass=Amazon.in.EHS'
            }
        }
    }
}
