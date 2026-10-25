# Keep serializable DTOs used by Retrofit with kotlinx.serialization.
-keep,includedescriptorclasses class com.example.outlook.data.**$$serializer { *; }
-keepclassmembers class com.example.outlook.data.** {
    *** Companion;
}
-keepclasseswithmembers class com.example.outlook.data.** {
    kotlinx.serialization.KSerializer serializer(...);
}
