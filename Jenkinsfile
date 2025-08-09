pipeline {
    agent any

    triggers {
        cron('0 5 * * *') // Run every day at 5 AM
    }

    stages {
        stage('Checkout Code') {
            steps {
                git branch: 'master', url: 'https://github.com/ansul9052/AmazonData.git'
            }
        }
        stage('Build') {
            steps {
                echo 'Building the project using Maven...'
                sh 'mvn clean compile'
            }
        }
        stage('Run Cab_Data') {
            steps {
                echo 'Running Cab_Data class...'
                sh 'mvn exec:java -Dexec.mainClass=Amazon.in.Cab_Data'
            }
        }
        stage('Run EHS') {
            steps {
                echo 'Running EHS class...'
                sh 'mvn exec:java -Dexec.mainClass=Amazon.in.EHS'
            }
        }
    }
}
