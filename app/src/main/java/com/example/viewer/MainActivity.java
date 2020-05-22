package com.example.viewer;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import android.Manifest;
import android.content.ContentResolver;
import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.text.method.ScrollingMovementMethod;
import android.util.Log;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.webkit.MimeTypeMap;
import android.widget.ImageView;
import android.widget.TextView;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xslf.extractor.XSLFPowerPointExtractor;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.hslf.model.Slide;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    static {
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLInputFactory",
                "com.fasterxml.aalto.stax.InputFactoryImpl"
        );
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLOutputFactory",
                "com.fasterxml.aalto.stax.OutputFactoryImpl"
        );
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLEventFactory",
                "com.fasterxml.aalto.stax.EventFactoryImpl"
        );
    }

    TextView textView;
    ImageView img_generalPhotos1,img_generalPhotos2,img_generalPhotos3,img_generalPhotos4,img_generalPhotos5,img_generalPhotos6,img_generalPhotos7,img_generalPhotos8,img_generalPhotos9,img_generalPhotos10,img_generalPhotos11,img_generalPhotos12;

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        MenuInflater inflater = getMenuInflater();
        inflater.inflate(R.menu.menu,menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        switch(item.getItemId()){
            case R.id.item:
                openDocumentFromFileManager();
                return true;
//            case R.id.item1:
//                return true;
            default:
                return super.onOptionsItemSelected(item);
        }
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        textView = (TextView)findViewById(R.id.textView);
    }

    private void openDocumentFromFileManager() {
        Intent i =new Intent();
        i.setType("application/*");
        i.setAction(Intent.ACTION_GET_CONTENT);
        if(PermissionHelper.getPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE,R.string.title_storage_permission,R.string.text_storage_permission,1111)){
            startActivityForResult(Intent.createChooser(i,"select document"),111);
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        try{
            if(resultCode == RESULT_OK){
                switch(requestCode){
                    case 111:
                        FileInputStream inputStream = (FileInputStream) getContentResolver().openInputStream(data.getData());

                        Uri uri = data.getData();
                        String extension = getMimeType(this,uri);
                        //Log.i("Extension",extension);

                        if(extension == "docx"){
                            XWPFDocument document = new XWPFDocument(inputStream);
                            XWPFWordExtractor extractor =new XWPFWordExtractor(document);

                            textView.setText(extractor.getText());
                            textView.setMovementMethod(new ScrollingMovementMethod());

                            int pages = document.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();
                            Log.i("Pages",Integer.toString(pages));
                        }
                        else if(extension == "doc"){

                            HWPFDocument document = new HWPFDocument(inputStream);
                            WordExtractor extractor = new WordExtractor(document);

                            textView.setText(extractor.getText());
                            textView.setMovementMethod(new ScrollingMovementMethod());

                            int pages = document.getSummaryInformation().getPageCount();
                            Log.i("Pages", Integer.toString(pages));
                        }
                        else if(extension == "pptx"){

                            XMLSlideShow slideShow = new XMLSlideShow(inputStream);

                            XSLFPowerPointExtractor extractor = new XSLFPowerPointExtractor(slideShow);

                            textView.setText(extractor.getText());
                            textView.setMovementMethod(new ScrollingMovementMethod());
//
                            int slides = slideShow.getSlides().length;
                            Log.i("Slides",Integer.toString(slides));

//                            try {
//
//                                FileInputStream fis = new FileInputStream("C:\\sample\\sample.pptx");
//                                XMLSlideShow pptxshow = new XMLSlideShow(fis);
//
//                                XSLFSlide[] slide2 = pptxshow.getSlides();
//                                for (int i = 0; i < slide2.length; i++) {
//                                    System.out.println(i);
//                                    try {
//                                        XSLFNotes mynotes = slide2[i].getNotes();
//                                        for (XSLFShape shape : mynotes) {
//                                            if (shape instanceof XSLFTextShape) {
//                                                XSLFTextShape txShape = (XSLFTextShape) shape;
//                                                for (XSLFTextParagraph xslfParagraph : txShape.getTextParagraphs()) {
//                                                    System.out.println(xslfParagraph.getText());
//
//                                                    textView.append(xslfParagraph.getText());
//                                                    textView.setMovementMethod(new ScrollingMovementMethod());
//                                                }
//                                            }
//                                        }
//                                    } catch (Exception e) {
//
//                                    }
//
//                                }
//                            } catch (IOException e) {
//
//                            }
                        }
                        else if(extension == "ppt"){

                            HSLFSlideShow show = new HSLFSlideShow(inputStream);
                            SlideShow ss = new SlideShow(show);
                            Slide[] slides = ss.getSlides();
                            for(int i=0;i< slides.length;i++){
                                TextRun[] runs = slides[i].getTextRuns();
                                for(int j=0;j<runs.length;j++){
                                    TextRun run = runs[j];
                                    if(run != null){
                                        String text = run.getText();
                                        textView.append(text);
                                        textView.setMovementMethod(new ScrollingMovementMethod());
                                    }
                                }
                            }

                            int slide_count = ss.getSlides().length;
                            Log.i("Slides",Integer.toString(slide_count));
                        }

                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String getMimeType(Context context, Uri uri) {
        String extension;

        if (uri.getScheme().equals(ContentResolver.SCHEME_CONTENT)) {
            final MimeTypeMap mime = MimeTypeMap.getSingleton();
            extension = mime.getExtensionFromMimeType(context.getContentResolver().getType(uri));
        } else {
            extension = MimeTypeMap.getFileExtensionFromUrl(Uri.fromFile(new File(uri.getPath())).toString());

        }

        return extension;
    }
}
