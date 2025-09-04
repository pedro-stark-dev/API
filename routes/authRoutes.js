const express = require('express')
const router = express.Router()
const authController = require('../controllers/authController')
const validateRegistro =  require('../middlewares/validateRegistro')
const { autenticarJWT } = require('../middlewares/authMiddleware');


router.post('/register', autenticarJWT, validateRegistro, authController.register);
router.post('/login',authController.login)
router.post('/logout', authController.logout)
router.post('/token', authController.renovarToken)
module.exports = router