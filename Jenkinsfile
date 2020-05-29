#!groovy
def BN = BRANCH_NAME == "master" || BRANCH_NAME.startsWith("releases/") ? BRANCH_NAME : "master"

library "knime-pipeline@$BN"

properties([
    pipelineTriggers([
        upstream('knime-base/' + env.BRANCH_NAME.replaceAll('/', '%2F'))
	]),
    buildDiscarder(logRotator(numToKeepStr: '5')),
    disableConcurrentBuilds()
])

SSHD_IMAGE = "${dockerTools.ECR}/knime/sshd:alpine3.10"

try {
    knimetools.defaultTychoBuild('org.knime.update.ext.poi')

    workflowTests.runTests(
        dependencies: [
            repositories: ["knime-excel", "knime-timeseries", "knime-jep", "knime-datageneration",
            "knime-filehandling", "knime-jfreechart", "knime-distance"]
        ],
        sidecarContainers: [
            [ image: SSHD_IMAGE, namePrefix: "SSHD", port: 22 ] 
        ]
    )

    stage('Sonarqube analysis') {
        env.lastStage = env.STAGE_NAME
        workflowTests.runSonar()
	}
} catch (ex) {
    currentBuild.result = 'FAILURE'
    throw ex
} finally {
    notifications.notifyBuild(currentBuild.result);
}
/* vim: set shiftwidth=4 expandtab smarttab: */
