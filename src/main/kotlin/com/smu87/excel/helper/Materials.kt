package com.smu87.excel.helper

data class Material(
    val supplier: String,
    val name: String,
    val units: String,
    val price: Double
) : Comparable<Material> {

    val isFromMainSupplier: Boolean = !SECONDARY_SUPPLIER.matches(supplier)

    override fun compareTo(other: Material): Int {
        // materials from main suppliers first
        val result = other.isFromMainSupplier.compareTo(isFromMainSupplier)
        if (result != 0) return result
        return name.compareTo(other.name)
    }

    companion object {
        private val SECONDARY_SUPPLIER: Regex = """\d*-\d*""".toRegex()
    }
}

data class MaterialInfo(
    val material: Material,
    val amount: Double
) {
    val cost: Double get() = amount * material.price

    operator fun plus(rhs: MaterialInfo): MaterialInfo {
        require(material == rhs.material)
        return MaterialInfo(material, amount + rhs.amount)
    }
}

class MaterialInfoBuilder {
    var supplier: String? = null
    var name: String? = null
    var units: String? = null
    var amount: Double? = null
    var price: Double? = null

    fun build(): MaterialInfo {
        val material = Material(supplier = supplier!!, name = name!!, units = units!!, price = price!!)
        return MaterialInfo(material, amount!!)
    }
}
