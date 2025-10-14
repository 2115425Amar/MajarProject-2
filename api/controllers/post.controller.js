// D:\WEBDEV\MajorProject\api\controllers\post.controller.js
import JWT from "jsonwebtoken";
import prisma from "../lib/prisma.js"; //model import
import fs from "fs";
import path from "path";
import XLSX from "xlsx";
import { generateDescription } from "../lib/gemini.js";

export const getPosts = async (req, res) => {
  const query = req.query;
  // console.log(query);
  try {
    const posts = await prisma.post.findMany({
      where: {
        city: query.city || undefined,
        type: query.type || undefined,
        property: query.property || undefined,
        property: query.property || undefined,
        bedroom: parseInt(query.bedroom) || undefined,
        price: {
          gte: parseInt(query.minPrice) || undefined,
          lte: parseInt(query.maxPrice) || undefined,
        },
      },
    });
    res.status(200).json(posts);
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: "Failed to get posts" });
  }
};

// ------------------------------------------------
const verifyToken = (token) => {
  // console.log("Verifying token:", token);
  if (!token) return null; // If no token is provided, return null
  return new Promise((resolve) => {
    JWT.verify(token, process.env.JWT_SECRET_KEY, (err, payload) => {
      if (err) return resolve(null); // If token is invalid or expired, resolve as null
      resolve(payload.id); // If token is valid, resolve with the user ID
    });
  });
};

export const getPost = async (req, res) => {
  const id = req.params.id;
  try {
    const post = await prisma.post.findUnique({
      where: { id },
      include: {
        postDetail: true,
        user: {
          select: {
            username: true,
            avatar: true,
          },
        },
      },
    });

    let userId = null;
    const token = req.cookies?.token;

    if (token) {
      userId = await verifyToken(token);
    }

    let saved = null;

    // Only check if the post is saved if the userId is not null
    if (userId) {
      saved = await prisma.savedPost.findUnique({
        where: {
          userId_postId: {
            postId: id,
            userId: userId,
          },
        },
      });
    }
    res.status(200).json({ ...post, isSaved: !!saved });
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: "Failed to get post" });
  }
};

// --------------------------------------------------

export const addPost = async (req, res) => {
  const body = req.body;
  const tokenUserId = req.userId;

  // console.log("Request Body:", body);
  // console.log("Formatted Data:", {
  //   ...body.postData,
  //   userId: tokenUserId,
  //   postDetail: {
  //     create: body.postDetail,
  //   },
  // });

  try {
    const newPost = await prisma.post.create({
      data: {
        ...body.postData,
        userId: tokenUserId,
        postDetail: {
          create: body.postDetail,
        },
      },
    });
    res.status(200).json(newPost);
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: "Failed to create post", err });
  }
};

// ------------------------------------------------------------------------

export const updatePost = async (req, res) => {
  try {
    const postId = req.params.id;
    const tokenUserId = req.userId;

    const existingPost = await prisma.post.findUnique({
      where: { id: postId },
    });

    if (!existingPost) {
      return res.status(404).json({ message: "Post not found" });
    }

    if (existingPost.userId !== tokenUserId) {
      return res.status(404).json({ message: "Not authorized" });
    }

    const updated = await prisma.post.update({
      where: { id: postId },
      data: {
        title: req.body.postData.title,
        price: req.body.postData.price,
        address: req.body.postData.address,
        city: req.body.postData.city,
        bedroom: req.body.postData.bedroom,
        bathroom: req.body.postData.bathroom,
        type: req.body.postData.type,
        property: req.body.postData.property,
        latitude: req.body.postData.latitude,
        longitude: req.body.postData.longitude,
        images: req.body.postData.images,
        postDetail: {
          update: {
            desc: req.body.postDetail.desc,
            utilities: req.body.postDetail.utilities,
            pet: req.body.postDetail.pet,
            income: req.body.postDetail.income,
            size: req.body.postDetail.size,
            school: req.body.postDetail.school,
            bus: req.body.postDetail.bus,
            restaurant: req.body.postDetail.restaurant,
          },
        },
      },
    });
    res.status(200).json(updated);
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: "Failed to update post" });
  }
};

export const deletePost = async (req, res) => {
  const id = req.params.id;
  const tokenUserId = req.userId;
  try {
    const post = await prisma.post.findUnique({
      where: { id },
    });

    if (post.userId !== tokenUserId) {
      return res.status(403).json({ message: "Not Authorized!" });
    }

    await prisma.post.delete({
      where: { id },
    });

    res.status(200).json({ message: "Post deleted" });
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: "Failed to delete post" });
  }
};

const categorizePrice = (price) => {
  if (price < 500000) return "Budget";
  if (price < 1500000) return "Mid-range";
  return "Luxury";
};

export const bulkUploadPosts = async (req, res) => {
  let filePath;
  try {
    if (!req.file) return res.status(400).json({ message: "No file uploaded" });

    filePath = path.join(process.cwd(), req.file.path);
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // req.file.path → location of temp file in uploads/ folder.
    // xlsx library opens the file (supports .csv & .xlsx).
    // SheetNames[0] gets the first sheet in Excel.
    // sheet_to_json converts the Excel/CSV rows into an array of JavaScript objects:

    const tokenUserId = req.userId; // from auth middleware

    // Maps each row to your Post model structure.
    // Splits Images column into an array of image URLs.
    // Converts numbers from strings.
    // Uses categorizePrice to classify into "Budget", "Mid-range", "Luxury".
    // Adds the logged-in user’s userId from token.

    const postsData = await Promise.all(
      jsonData.map(async (row) => {
        // Generate description if missing
        let description = row.description;
        if (!description || description.trim() === "") {
          description = await generateDescription(row);
        }
        return {
          title: row.title,
          price: parseInt(row.price, 10),
          images: row.images ? row.images.split(",") : [],
          address: row.address,
          city: row.city,
          bedroom: parseInt(row.bedroom, 10),
          bathroom: parseInt(row.bathroom, 10),
          latitude: String(row.latitude),
          longitude: String(row.longitude),
          type: row.type?.toLowerCase(),
          property: row.property?.toLowerCase(),
          postDetail: {
            create: {
              // desc: row.description || "",
              desc: description,
              size: parseInt(row.size, 10) || null,
              school: row.school || null,
              bus: row.bus || null,
              restaurant: row.restaurant || null,
            },
          },
          category: categorizePrice(parseInt(row.price, 10)),
          userId: tokenUserId,
        };
      })
    );

    // prisma.$transaction executes all inserts in one DB transaction.
    // If any insert fails → the whole batch is rolled back.
    await prisma.$transaction(
      postsData.map((post) =>
        prisma.post.create({
          data: post,
        })
      )
    );
    // fs.unlinkSync(filePath); // cleanup
    res.status(200).json({ message: `${postsData.length} posts uploaded successfully` });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: "Bulk upload failed" });
  } finally {
    if (filePath && fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  }
};

///// for recommendations /////
export const getPostsByIds = async (req, res) => {
  try {
    const ids = req.query.ids?.split(",") || [];
    // Validate IDs: only keep those with length 24 (MongoDB ObjectId)
    const validIds = ids.filter(id => id.length === 24);
    if (validIds.length === 0) return res.json([]);

    const posts = await prisma.post.findMany({
      where: { id: { in: validIds } }
    });
    res.json(posts);
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: "Failed to fetch posts by IDs" });
  }
};