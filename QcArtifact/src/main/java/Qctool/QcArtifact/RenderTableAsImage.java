package Qctool.QcArtifact;


import com.aspose.words.*;
import com.aspose.words.Shape;


import javax.imageio.ImageIO;
import java.awt.*;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;

import static com.aspose.words.NodeType.*;

public class RenderTableAsImage {

// Logic for cloning the document and then rendering
    public static void render_Node(Node node, ImageSaveOptions imageOptions) throws Exception {
        if (node == null) throw new Exception("Node cannot be null");
        // If no image options are supplied, create default options.
        if (imageOptions == null) imageOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageOptions.setPaperColor(new Color(0, 0, 0, 0));
        // There a bug which affects the cache of a cloned node. To avoid this we instead clone the entire document including all nodes, // find the matching node in the cloned document and render that instead.
        Document doc = ((Document) node.getDocument()).deepClone();
        node = doc.getChild(TABLE, node.getDocument().getChildNodes(TABLE, true).indexOf(node), true);
        // Create a temporary shape to store the target node in. This shape will be rendered to retrieve // the rendered content of the node.
        Shape shape = new Shape(doc, ShapeType.TEXT_BOX);
        Section parentSection = (Section) node.getAncestor(SECTION);
        // Assume that the node cannot be larger than the page in size.
        shape.setWidth(parentSection.getPageSetup().getPageWidth());
        shape.setHeight(parentSection.getPageSetup().getPageHeight());
        shape.setFillColor(new Color(255,255,255));
        // We must make the shape and paper color transparent. // Don't draw a surrounding line on the
        shape.setStroked(false);
        Node currentNode = node;
        // If the node contains block level nodes then just add a copy of these nodes to the shape.
        if (currentNode instanceof InlineStory || currentNode instanceof Story) {
            CompositeNode<Node> composite = (CompositeNode) currentNode;
            for (Node childNode : (Iterable<Node>) composite.getChildNodes(TABLE, true)) {
                shape.appendChild(childNode.deepClone(true));
            }
        } else { // Move up through the DOM until we find node which is suitable to insert into a Shape (a node with a parent can contain paragraph, tables the same as a shape). // Each parent node is cloned on the way up so even a descendant node passed to this method can be rendered. // Since we are working with the actual nodes of the document we need to clone the target node into the temporary shape.
            while (!(currentNode.getParentNode() instanceof InlineStory || currentNode.getParentNode() instanceof Story || currentNode.getParentNode() instanceof ShapeBase || currentNode.getNodeType() == PARAGRAPH)) {
                CompositeNode parent = (CompositeNode) currentNode.getParentNode().deepClone(false);
                currentNode = currentNode.getParentNode();
                parent.appendChild(node.deepClone(true));
                node = parent;}
                // Store this new node to be inserted into the shape.
                //
                // Add the node to the shape.
                shape.appendChild(node.deepClone(true));
            } // We must add the shape to the document tree to have it rendered.
            parentSection.getBody().getFirstParagraph().appendChild(shape);
            shape.getShapeRenderer().save("C:\\Users\\tgaur\\OneDrive\\Desktop\\QcArtifact\\QcArtifact\\src\\main\\resources" + "Out.png", imageOptions);

        }


    public static Rectangle FindBoundingBoxAroundNode(BufferedImage originalBitmap)
    {
        Point min = new Point(Integer.MAX_VALUE, Integer.MAX_VALUE);
        Point max = new Point(Integer.MIN_VALUE, Integer.MIN_VALUE);
        for (int x = 0; x <originalBitmap.getWidth(); ++x)
        {
            for (int y = 0; y <originalBitmap.getHeight(); ++y)
            {
                int argb = originalBitmap.getRGB(x, y);
                if (argb != new Color(255,255,255).getRGB())
                {
                    min.x = Math.min(x, min.x);
                    min.y = Math.min(y, min.y);
                    max.x = Math.max(x, max.x);
                    max.y = Math.max(y, max.y);

                }
            }
        }

        return new Rectangle(min.x, min.y, (max.x - min.x) + 1, (max.y - min.y) + 1);
    }

    public static void main(String[] args) throws Exception {
        String a = "C:\\Users\\tgaur\\OneDrive\\Desktop\\QcArtifact\\QcArtifact\\src\\main\\table.docx";
        Document doc = new Document(a);
        NodeCollection tables = doc.getChildNodes(TABLE, true);

            Table table = (Table) tables.get(0);
           // render_Table(table, new ImageSaveOptions(SaveFormat.PNG));
           /* BufferedImage renderedImage = ImageIO.read(new File("C:\\Users\\tgaur\\OneDrive\\Desktop\\QcArtifact\\QcArtifact\\src\\main\\resources" + "Out.png"));
            // Extract the actual content of the image by cropping transparent space around // the rendered shape.
            Rectangle cropRectangle = FindBoundingBoxAroundNode(renderedImage);
            BufferedImage out = renderedImage.getSubimage(cropRectangle.x, cropRectangle.y, cropRectangle.width, cropRectangle.height);
            File outputfile = new File("C:\\Users\\tgaur\\OneDrive\\Desktop\\QcArtifact\\QcArtifact\\src\\main\\resources235"+ "Out.png");
            ImageIO.write(out, "png", outputfile);*/

        byte[] imageBytes = render_Tablebytes(table, new ImageSaveOptions(SaveFormat.PNG));

        // Create BufferedImage from the byte array
        BufferedImage renderedImage = createImageFromBytes(imageBytes);

        // Extract the actual content of the image by cropping transparent space around the rendered shape.
        Rectangle cropRectangle = FindBoundingBoxAroundNode(renderedImage);
        BufferedImage out = renderedImage.getSubimage(cropRectangle.x, cropRectangle.y, cropRectangle.width, cropRectangle.height);

        File outputfile = new File("C:\\Users\\tgaur\\OneDrive\\Desktop\\QcArtifact\\QcArtifact\\src\\main\\resources235" + "Out.png");
        ImageIO.write(out, "png", outputfile);
            System.out.println("Parsing over" );


        System.out.println("parsing completed");
    }

    // deals only with table Node
    public static void render_Table(Table table, ImageSaveOptions imageOptions) throws Exception {
        if (table == null) throw new Exception("Table cannot be null");
        // If no image options are supplied, create default options.
        if (imageOptions == null) imageOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageOptions.setPaperColor(new Color(0, 0, 0, 0));

        // Create a temporary shape to store the target node in.
        Shape shape = new Shape(table.getDocument(), ShapeType.TEXT_BOX);
        Section parentSection = (Section) table.getAncestor(SECTION);

        // Assume that the table cannot be larger than the page in size.
        shape.setWidth(parentSection.getPageSetup().getPageWidth());
        shape.setHeight(parentSection.getPageSetup().getPageHeight());
        shape.setFillColor(new Color(255, 255, 255));
        shape.setStroked(false);

        // Add a copy of the table to the shape.
        shape.appendChild(table.deepClone(true));

        // Add the shape to the document tree to have it rendered.
        parentSection.getBody().getFirstParagraph().appendChild(shape);

        // Save the rendered shape as an image.
        shape.getShapeRenderer().save("C:\\Users\\tgaur\\OneDrive\\Desktop\\QcArtifact\\QcArtifact\\src\\main\\resources" + "Out.png", imageOptions);
    }


        public static byte[] render_Tablebytes(Table table, ImageSaveOptions imageOptions) throws Exception {
            if (table == null) throw new Exception("Table cannot be null");
            // If no image options are supplied, create default options.
            if (imageOptions == null) imageOptions = new ImageSaveOptions(SaveFormat.PNG);
            imageOptions.setPaperColor(new Color(0, 0, 0, 0));

            // Create a temporary shape to store the target node in.
            Shape shape = new Shape(table.getDocument(), ShapeType.TEXT_BOX);
            Section parentSection = (Section) table.getAncestor(SECTION);

            // Assume that the table cannot be larger than the page in size.
            shape.setWidth(parentSection.getPageSetup().getPageWidth());
            shape.setHeight(parentSection.getPageSetup().getPageHeight());
            shape.setFillColor(new Color(255, 255, 255));
            shape.setStroked(false);

            // Add a copy of the table to the shape.
            shape.appendChild(table.deepClone(true));

            // Add the shape to the document tree to have it rendered.
            parentSection.getBody().getFirstParagraph().appendChild(shape);

            // Save the rendered shape as an image in memory.
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            shape.getShapeRenderer().save(outputStream, imageOptions);

            // Return the byte array of the rendered image.
            return outputStream.toByteArray();
        }
    public static BufferedImage createImageFromBytes(byte[] imageBytes) throws IOException {
        ByteArrayInputStream inputStream = new ByteArrayInputStream(imageBytes);
        return ImageIO.read(inputStream);
    }
    }





