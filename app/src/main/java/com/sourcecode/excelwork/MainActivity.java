package com.sourcecode.excelwork;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Iterator;
import java.util.StringTokenizer;

import javax.mail.AuthenticationFailedException;
import javax.mail.MessagingException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import android.app.Activity;
import android.app.NotificationManager;
import android.content.Context;
import android.content.Intent;
import android.graphics.Typeface;
import android.net.Uri;
import android.os.AsyncTask;
import android.os.Bundle;
import android.os.Environment;
import android.os.StrictMode;
import android.support.v4.app.NotificationCompat;
import android.util.Log;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

public class MainActivity extends Activity {
	 GMailSender sender;
	EditText et1,et2;
	TextView tv1,tv2,tv3;
	String path = null,sender_email;
	static int count=0;
	String emailadd=null;
	private static final int FILE_SELECT_CODE = 0;
	private static int flag=0;
	NotificationManager mNotifyMgr;
    NotificationCompat.Builder mBuilder;
    int mNotificationId;
	
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		if (android.os.Build.VERSION.SDK_INT > 9) {
		    StrictMode.ThreadPolicy policy = new StrictMode.ThreadPolicy.Builder().permitAll().build();
		    StrictMode.setThreadPolicy(policy);
		}
		et1=(EditText)findViewById(R.id.editText1);
		et2=(EditText)findViewById(R.id.editText2);
		tv1=(TextView)findViewById(R.id.textView3);	
		tv2=(TextView)findViewById(R.id.textView4);
		tv3=(TextView)findViewById(R.id.textView1);
		
		Typeface tf=Typeface.createFromAsset(getAssets(),"fonts/Pacifico.ttf");
        tv3.setTypeface(tf);
        tf=Typeface.createFromAsset(getAssets(), "fonts/RobotoCondensed-Light.ttf");
        tv1.setTypeface(tf);
        tv2.setTypeface(tf);
        
		
	}

	public void onFileExcelSelect(View v){
		Intent intent = new Intent(Intent.ACTION_GET_CONTENT); 
	    intent.setType("*/*"); 
	    intent.addCategory(Intent.CATEGORY_OPENABLE);

	    try {
	    	flag=1;
	        startActivityForResult(
	                Intent.createChooser(intent, "Select a File to Upload"),
	                FILE_SELECT_CODE);
	    } catch (android.content.ActivityNotFoundException ex) {
	        Toast.makeText(this, "Please install a File Manager.", 
	                Toast.LENGTH_SHORT).show();
	    }
	}
	
	public void onFileHtmlSelect(View v){
		Intent intent = new Intent(Intent.ACTION_GET_CONTENT); 
	    intent.setType("*/*"); 
	    intent.addCategory(Intent.CATEGORY_OPENABLE);
	    try {
	    	flag=0;
	        startActivityForResult(
	                Intent.createChooser(intent, "Select a File to Upload"),
	                FILE_SELECT_CODE);
	    } catch (android.content.ActivityNotFoundException ex) {
	        Toast.makeText(this, "Please install a File Manager.", 
	                Toast.LENGTH_SHORT).show();
	    }
	}
	
	@Override
	protected void onActivityResult(int requestCode, int resultCode, Intent data) {
	    switch (requestCode) {
	        case FILE_SELECT_CODE:
	        if (resultCode == RESULT_OK) {
	            Uri uri = data.getData();
	           	path = FileUtils.getPath(getApplicationContext(), uri);
	            if(flag==1)
	            {	
	            	tv1.setText(path);
	            }
	            else
	            {
	            	tv2.setText(path);
	            }
	           }
	        break;
	    }
	    super.onActivityResult(requestCode, resultCode, data);
	}
	
	public void onSend(View v) {
		sender_email=et1.getText().toString();
		sender = new GMailSender(sender_email,et2.getText().toString());
		StringBuilder contentBuilder = new StringBuilder();
		try {
		    BufferedReader in = new BufferedReader(new FileReader(tv2.getText().toString()));
		    String str;
		    while ((str = in.readLine()) != null) {
		        contentBuilder.append(str);
		    }
		    in.close();
		} catch (IOException e) {
		}
		String content = contentBuilder.toString();
		if(path!=null)
		{
			emailadd=readExcelFile(MainActivity.this,tv1.getText().toString());	
			mNotifyMgr = (NotificationManager) getSystemService(NOTIFICATION_SERVICE);
	        mBuilder= new NotificationCompat.Builder(getApplicationContext())
	        .setSmallIcon(R.drawable.ic_launcher)
	        .setContentTitle("Sending Emails")
	        .setContentText("Sending in progress")
	        .setAutoCancel(true);
	        mNotificationId = 001;
				new SendEmailAsyncTask().execute(emailadd,content);	
		}
	}
	
	public class SendEmailAsyncTask extends AsyncTask <String, Integer, Boolean> {
				
		private int count=0,fcount=0;
				
		public SendEmailAsyncTask() {
                      }
                
		                @Override
						protected void onProgressUpdate(Integer... values) {
		                	mBuilder.setProgress(100,values[0], false);
            				mNotifyMgr.notify(mNotificationId, mBuilder.build());
							super.onProgressUpdate(values);
						}

						@Override
						protected void onPreExecute() {
							mBuilder.setProgress(100, 0, false);
							mNotifyMgr.notify(mNotificationId, mBuilder.build());
							super.onPreExecute();
						}

						@Override
						protected void onPostExecute(Boolean result) {
							//mBuilder.setContentTitle("Sending Complete");
							mBuilder.setContentText("Total Email Sent: "+count+"\n Failed: "+fcount);
							mBuilder.setProgress(0, 0, false);
							mNotifyMgr.notify(mNotificationId, mBuilder.build());
							super.onPostExecute(result);
						}

						@Override
                protected Boolean doInBackground(String... params) {
                    try {
                    	StringTokenizer st2 = new StringTokenizer(params[0], ",");
            			
            			while (st2.hasMoreElements()) {
                            Log.d("email address",st2.nextElement().toString());
            				sender.sendMail("Source Code Info Tech Pvt. Ltd. Invitation to Free download Study Material, Final year Projects and Research-Review Papers for Student", params[1],
                                    sender_email,st2.nextElement().toString());	
            				count++;
            				publishProgress(Math.min((int)((float)(100*count)/900), 100));
            			}	
                        return true;
                    } catch (AuthenticationFailedException e) {
                        fcount++;
                        return false;
                    } catch (MessagingException e) {
                        fcount++;
                        return false;
                    } catch (Exception e) {
                        fcount++;
                        return false;
                    }
                }
         }   
         
         private static String readExcelFile(Context c,String filename) { 
        	 String emailadd ="";
             if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) 
             {
				 return emailadd;
             }

             try{
                 FileInputStream myInput = new FileInputStream(filename);
                 POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);
                 HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);
                 HSSFSheet mySheet = myWorkBook.getSheetAt(0);
                 Iterator<Row> rowIter = mySheet.rowIterator();
                
                 while(rowIter.hasNext()){
                	 if(count==900)
                	 {
                		 break;
                	 }
                	 else{
                		 	HSSFRow myRow = (HSSFRow) rowIter.next();
                		 	Iterator<Cell> cellIter = myRow.cellIterator();
                		 	count++;
                		 	if(cellIter.hasNext()){
                		 		HSSFCell myCell = (HSSFCell) cellIter.next();
                		 		emailadd+=myCell.toString()+",";
							}
                	 }
                }
             }catch (Exception e){e.printStackTrace(); }

             return emailadd;
         }
         
         public static boolean isExternalStorageReadOnly() { 
             String extStorageState = Environment.getExternalStorageState(); 
             if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) { 
                 return true; 
             } 
             return false; 
         } 

         public static boolean isExternalStorageAvailable() { 
             String extStorageState = Environment.getExternalStorageState(); 
             if (Environment.MEDIA_MOUNTED.equals(extStorageState)) { 
                 return true; 
             } 
             return false; 
         } 
         
	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		getMenuInflater().inflate(R.menu.main, menu);
		return true;
	}

	@Override
	public boolean onOptionsItemSelected(MenuItem item) {
		int id = item.getItemId();
		if (id == R.id.action_settings) {
			return true;
		}
		return super.onOptionsItemSelected(item);
	}
}
