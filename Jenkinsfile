pipeline {
  agent any
  stages {
    stage('Stage') {
      steps {
        bat(script: 'permisos.ps1', encoding: 'powershell', label: 'Permisos', returnStatus: true, returnStdout: true)
      }
    }

  }
}