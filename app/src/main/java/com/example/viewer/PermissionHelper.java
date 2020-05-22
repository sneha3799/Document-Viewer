package com.example.viewer;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.DialogInterface;
import android.content.pm.PackageManager;

import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;

public class PermissionHelper {
    public static boolean getPermission(final Activity activity, final String permission, int title, int message, final int id){
        if(ContextCompat.checkSelfPermission(activity,permission) != PackageManager.PERMISSION_GRANTED){
            if(ActivityCompat.shouldShowRequestPermissionRationale(activity,permission)){
                new AlertDialog.Builder(activity)
                                .setTitle(title)
                                .setMessage(message)
                                .setPositiveButton("ok", new DialogInterface.OnClickListener() {
                                    @Override
                                    public void onClick(DialogInterface dialog, int which) {
                                        ActivityCompat.requestPermissions(activity, new String[] {permission},id);
                                    }
                                })
                                .create()
                                .show();
            }
            else{
                ActivityCompat.requestPermissions(activity, new String[] {permission},id);
            }
            return false;
        }else{
            return true;
        }
    }
}
