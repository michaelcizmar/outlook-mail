import java.util.Properties

plugins {
    alias(libs.plugins.android.application)
}

// The signature hash goes into the manifest intent filter for MSAL's redirect.
// Set msal.signatureHash in local.properties (not committed). See README.
val localProps = Properties().apply {
    val file = rootProject.file("local.properties")
    if (file.exists()) file.inputStream().use { load(it) }
}
val msalSignatureHash: String = localProps.getProperty("msal.signatureHash") ?: "REPLACE_WITH_SIGNATURE_HASH"

android {
    namespace = "com.example.outlook.harness"
    compileSdk = 37

    defaultConfig {
        applicationId = "com.example.outlook.harness"
        minSdk = 26
        targetSdk = 37
        versionCode = 1
        versionName = "0.1.0"
        manifestPlaceholders["msalSignatureHash"] = msalSignatureHash
    }

    buildTypes {
        release {
            isMinifyEnabled = false
        }
    }

    lint {
        // The harness is intentionally unstyled and untranslated.
        disable += setOf("SetTextI18n", "MissingApplicationIcon", "UseKtx")
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
}

dependencies {
    implementation(project(":outlook-lib"))
    implementation(libs.androidx.activity.ktx)
}
