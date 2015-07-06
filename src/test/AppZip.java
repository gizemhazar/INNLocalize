package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class AppZip {

	public void zipFolder(String zFilePath, String srcFolder, String destZipFile)
			throws Exception {
		ZipOutputStream zip = null;
		FileOutputStream fileWriter = null;
		File zFile = new File(zFilePath, destZipFile);
		fileWriter = new FileOutputStream(zFile);
		zip = new ZipOutputStream(fileWriter);

		addFolderToZip("", srcFolder, zip);
		zip.flush();
		zip.close();
	}

	private void addFileToZip(String path, String srcFile, ZipOutputStream zip)
			throws Exception {

		File folder = new File(srcFile);
		if (folder.isDirectory()) {
			addFolderToZip(path, srcFile, zip);
		} else {
			byte[] buf = new byte[1024];
			int len;
			FileInputStream in = new FileInputStream(srcFile);
			zip.putNextEntry(new ZipEntry(path + "/" + folder.getName()));
			while ((len = in.read(buf)) > 0) {
				zip.write(buf, 0, len);
			}
		}
	}

	private void addFolderToZip(String path, String srcFolder,
			ZipOutputStream zip) throws Exception {
		File folder = new File(srcFolder);

		for (String fileName : folder.list()) {
			if (path.equals("")) {
				addFileToZip(folder.getName(), srcFolder + "/" + fileName, zip);
			} else {
				addFileToZip(path + "/" + folder.getName(), srcFolder + "/"
						+ fileName, zip);
			}
		}
	}
}
