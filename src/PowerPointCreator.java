import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hslf.model.*;
import org.apache.poi.hslf.model.Shape;
import org.apache.poi.hslf.usermodel.RichTextRun;
import org.apache.poi.hslf.usermodel.SlideShow;

import java.awt.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

/**
 * This class creates EBCSV praise Powerpoints
 *
 * Usage: PowerPointCreator <last week's ppt> <1st song ppt> <2nd song ppt> <3rd song ppt> <4th song ppt>
 *
 * @author Eric Chang
 * @since 2013
 *
 */
public class PowerPointCreator {
    private static final String TMP_DIR = "tmp/";
    private static final String ANNOUNCEMENTS_PPT_STRING = "announcements.ppt";

    private static File createFileCopy(File file) throws IOException {
        File tmpFile = new File(TMP_DIR + file.getName());
        FileUtils.copyFile(file, tmpFile);
        return tmpFile;
    }

    private static void addAnnouncements(SlideShow slideShow) throws IOException {
        File file = new File(ANNOUNCEMENTS_PPT_STRING);
        OutputStream os = new FileOutputStream(file);
        InputStream is = PowerPointCreator.class.getResourceAsStream(ANNOUNCEMENTS_PPT_STRING);
        IOUtils.copy(is, os);
        os.close();

        SlideShow announcementsSlideShow = getSlideShow(file);
        Slide[] slides = announcementsSlideShow.getSlides();

        for (Slide slide : slides) {
            Slide newSlide = slideShow.createSlide();
            Shape[] shapes = slide.getShapes();
            for (Shape shape : shapes) {
                newSlide.addShape(shape);
            }
        }

        // clean up tmp announcements file
        FileUtils.deleteQuietly(file);
    }

    private static void addSongSlides(SlideShow slideShow, List<File> songFiles) throws IOException {
        for (File songFile : songFiles) {
            SlideShow songSlideShow = getSlideShow(songFile);
            Slide[] songSlides = songSlideShow.getSlides();
            for (Slide songSlide : songSlides) {
                addSongSlide(slideShow, songSlide);
            }
        }
    }

    private static void addSongSlide(SlideShow slideShow, Slide slide) throws IOException {
        Slide newSlide = slideShow.createSlide();

        // set slide title
        TextBox textBox = newSlide.addTitle();
        String songTitle = slide.getTitle();
        textBox.setText(songTitle);

        // set slide content
        Shape[] shapes = slide.getShapes();
        for (Shape shape : shapes) {
            if (shape instanceof AutoShape) {
                AutoShape autoShape = (AutoShape) shape;
                if (songTitle.equals(autoShape.getText())) {
                    continue;  // skip old title shape
                }
                autoShape.getFill().setFillType(Fill.FILL_BACKGROUND);
            }
            newSlide.addShape(shape);
        }

        // set title color
        TextRun textRun = newSlide.getTextRuns()[0];
        for (RichTextRun richTextRun : textRun.getRichTextRuns()) {
            richTextRun.setFontColor(Color.CYAN);
        }
    }

    private static SlideShow getSlideShow(File file) throws IOException {
        FileInputStream is = new FileInputStream(file);
        SlideShow slideShow = new SlideShow(is);
        is.close();
        return slideShow;
    }

    private static void saveSlideShow(SlideShow slideShow, File file) throws IOException {
        FileOutputStream out = new FileOutputStream(file);
        slideShow.write(out);
        out.close();
    }

    private static String getDropBoxPresentationFileName() {
        Calendar cal = Calendar.getInstance(TimeZone.getTimeZone("PST"));

        // find next thursday or sunday
        int dayOfWeek = cal.get(Calendar.DAY_OF_WEEK);
        int offset = 0;
        switch (dayOfWeek) {
            case Calendar.MONDAY:
                offset = 3;
                break;
            case Calendar.TUESDAY:
            case Calendar.FRIDAY:
                offset = 2;
                break;
            case Calendar.WEDNESDAY:
            case Calendar.SATURDAY:
                offset = 1;
                break;
        }
        cal.add(Calendar.DAY_OF_MONTH, offset);

        dayOfWeek = cal.get(Calendar.DAY_OF_WEEK);
        if (dayOfWeek == Calendar.SUNDAY) {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy_MM_dd");
            return "ebcsv_sunday_" + sdf.format(cal.getTime()) + ".ppt";
        } else {
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
            return "TBS_" + sdf.format(cal.getTime()) + ".ppt";
        }
    }

    private static void printUsage() {
        System.out.println("Usage: java -jar PowerPointCreator.jar <last week's ppt> <1st song ppt> <2nd song ppt> <3rd song ppt> <4th song ppt>");
    }

    public static void main (String[] args) throws Exception {
        if (args.length < 1) {
            printUsage();
            return;
        }

        File titleSlideDropBoxFile = new File(args[0]);
        File titleSlideFile = createFileCopy(titleSlideDropBoxFile);

        // copies files from input to prevent damaging original files in dropbox
        ArrayList<File> songFiles = new ArrayList<File>();
        for (int i = 1; i < args.length; i++) {
            File songFile = new File(args[i]);
            songFile = createFileCopy(songFile);
            songFiles.add(songFile);
        }

        SlideShow ppt = getSlideShow(titleSlideFile);
        int length = ppt.getSlides().length;

        // remove all slides except title slide
        for (int i=length-1; i != 0; i--) {
            ppt.removeSlide(i);
        }

        addSongSlides(ppt, songFiles);
        addAnnouncements(ppt);

        // save to dropbox
        File dropBoxPptFile = new File(titleSlideDropBoxFile.getParent() + File.separator + getDropBoxPresentationFileName());
        saveSlideShow(ppt, dropBoxPptFile);
        System.out.println("Successfully created " + dropBoxPptFile.getAbsolutePath());

        // clean tmp directory
        FileUtils.deleteDirectory(new File(TMP_DIR));
    }

}
